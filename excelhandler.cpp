#include "excelhandler.h"
#include "reportdatamodel.h"
#include "Cell.h"

#include <QMessageBox>
#include <QFileInfo>
#include <QApplication>
#include <QProgressDialog>
#include <memory> // for std::make_unique
#include <QSet>
#include <QTextCodec>
#include <qDebug>

// 明确包含所有需要的QXlsx库头文件
#include "xlsxdocument.h"
#include "xlsxworksheet.h"
#include "xlsxcell.h"
#include "xlsxformat.h"
#include "xlsxcellrange.h"


bool ExcelHandler::loadFromFile(const QString& fileName, ReportDataModel* model)
{
    if (fileName.isEmpty() || !model) {
        QMessageBox::warning(nullptr, "错误", "参数无效");
        return false;
    }
    if (!isValidExcelFile(fileName)) {
        return false;
    }

    auto progress = std::make_unique<QProgressDialog>("正在导入Excel文件...", "取消", 0, 100, nullptr);
    progress->setWindowModality(Qt::ApplicationModal);
    progress->setMinimumDuration(0);
    progress->show();

    model->clearAllCells();

    QXlsx::Document xlsx(fileName);
    if (!xlsx.load()) {
        QMessageBox::warning(nullptr, "错误", "无法打开Excel文件，文件可能已损坏或格式不支持");
        return false;
    }
    progress->setValue(20);
    qApp->processEvents();

    QXlsx::Worksheet* worksheet = static_cast<QXlsx::Worksheet*>(xlsx.currentSheet());
    if (!worksheet) {
        QMessageBox::warning(nullptr, "错误", "无法读取工作表内容");
        return false;
    }

    const QXlsx::CellRange range = worksheet->dimension();
    if (!range.isValid()) {
        model->updateModelSize(0, 0);
        return true;
    }

    int totalRows = range.rowCount();
    int maxModelRow = 0, maxModelCol = 0;

    // 首先处理合并单元格信息
    QHash<QPoint, RTMergedRange> mergedRanges;
    loadMergedCells(worksheet, mergedRanges);
    progress->setValue(30);
    qApp->processEvents();

    // 处理单元格数据
    for (int row = range.firstRow(); row <= range.lastRow(); ++row) {
        if (progress->wasCanceled()) break;
        for (int col = range.firstColumn(); col <= range.lastColumn(); ++col) {

            auto xlsxCell = worksheet->cellAt(row, col);

            if (xlsxCell) {
                RTCell* newCell = new RTCell();

                // 读取公式和值
                QVariant rawData = worksheet->read(row, col);
                if (rawData.type() == QVariant::String && rawData.toString().startsWith('=')) {
                    newCell->setFormula(rawData.toString());
                }
                newCell->value = xlsxCell->value();

                // 转换格式
                if (xlsxCell->format().isValid()) {
                    convertFromExcelStyle(xlsxCell->format(), newCell->style);
                }

                // 设置合并单元格信息
                QPoint cellPos(row - 1, col - 1);  // 转换为0基索引
                if (mergedRanges.contains(cellPos)) {
                    newCell->mergedRange = mergedRanges[cellPos];
                }

                int modelRow = row - 1, modelCol = col - 1;
                model->addCellDirect(modelRow, modelCol, newCell);
                maxModelRow = qMax(maxModelRow, modelRow + 1);
                maxModelCol = qMax(maxModelCol, modelCol + 1);
            }
        }
        int progressValue = 30 + ((row - range.firstRow() + 1) * 60 / totalRows);
        progress->setValue(progressValue);
        qApp->processEvents();
    }

    if (progress->wasCanceled()) {
        model->clearAllCells();
        model->updateModelSize(0, 0);
        return false;
    }

    model->updateModelSize(maxModelRow, maxModelCol);

    // 处理合并单元格的样式统一 - 只处理每个合并区域一次
    QSet<QPair<QPair<int, int>, QPair<int, int>>> processedMergedRanges;

    for (auto it = mergedRanges.constBegin(); it != mergedRanges.constEnd(); ++it) {
        const RTMergedRange& mergedRange = it.value();

        // 创建合并范围的唯一标识符
        QPair<QPair<int, int>, QPair<int, int>> rangeKey = {
            {mergedRange.startRow, mergedRange.startCol},
            {mergedRange.endRow, mergedRange.endCol}
        };

        // 如果这个合并区域已经处理过，跳过
        if (processedMergedRanges.contains(rangeKey)) {
            continue;
        }

        processedMergedRanges.insert(rangeKey);

        // 获取主单元格（左上角）
        RTCell* mainCell = model->getCell(mergedRange.startRow, mergedRange.startCol);
        if (!mainCell) {
            // 如果主单元格不存在，创建一个
            mainCell = model->ensureCell(mergedRange.startRow, mergedRange.startCol);
            mainCell->mergedRange = mergedRange;
        }

        // 将主单元格的样式复制到合并区域内的所有单元格
        for (int row = mergedRange.startRow; row <= mergedRange.endRow; ++row) {
            for (int col = mergedRange.startCol; col <= mergedRange.endCol; ++col) {
                RTCell* cell = model->ensureCell(row, col);
                if (cell) {
                    // 设置合并信息
                    cell->mergedRange = mergedRange;

                    // 如果不是主单元格，清空内容但保持样式
                    if (row != mergedRange.startRow || col != mergedRange.startCol) {
                        cell->value = QVariant();
                        cell->formula.clear();
                        cell->hasFormula = false;
                    }

                    // 应用主单元格的样式到所有单元格
                    cell->style = mainCell->style;
                }
            }
        }
    }

    progress->setValue(100);
    return true;
}

void ExcelHandler::loadRowColumnSizes(QXlsx::Worksheet* worksheet, ReportDataModel* model)
{
    // 读取列宽
    for (int col = 1; col <= worksheet->dimension().lastColumn(); ++col) {
        double width = worksheet->columnWidth(col);
        if (width > 0) {
            // 转换为像素并存储到模型中
            // 注意：QXlsx的宽度单位可能需要转换
            qDebug() << "Column" << col << "width:" << width;
        }
    }

    // 读取行高
    for (int row = 1; row <= worksheet->dimension().lastRow(); ++row) {
        double height = worksheet->rowHeight(row);
        if (height > 0) {
            qDebug() << "Row" << row << "height:" << height;
        }
    }
}

bool ExcelHandler::saveToFile(const QString& fileName, ReportDataModel* model)
{
    if (fileName.isEmpty() || !model) {
        QMessageBox::warning(nullptr, "错误", "参数无效");
        return false;
    }
    QString actualFileName = fileName;
    if (!actualFileName.endsWith(".xlsx", Qt::CaseInsensitive)) {
        actualFileName += ".xlsx";
    }

    auto progress = std::make_unique<QProgressDialog>("正在导出Excel文件...", "取消", 0, 100, nullptr);
    progress->setWindowModality(Qt::ApplicationModal);
    progress->setMinimumDuration(0);
    progress->show();

    QXlsx::Document xlsx;
    QXlsx::Worksheet* worksheet = xlsx.currentWorksheet();
    if (!worksheet) {
        QMessageBox::warning(nullptr, "错误", "无法创建Excel工作表");
        return false;
    }
    progress->setValue(10);

    const auto& allCells = model->getAllCells();
    int totalCells = allCells.size();
    int processedCells = 0;

    // 保存单元格数据和格式
    for (auto it = allCells.constBegin(); it != allCells.constEnd(); ++it) {
        if (progress->wasCanceled()) return false;
        const QPoint& modelPos = it.key();
        const RTCell* cell = it.value();

        int excelRow = modelPos.x() + 1;
        int excelCol = modelPos.y() + 1;
        QXlsx::Format cellFormat = convertToExcelFormat(cell->style);

        if (cell->hasFormula) {
            QString fullFormula = cell->formula.startsWith('=') ? cell->formula : ("=" + cell->formula);
            worksheet->write(excelRow, excelCol, fullFormula, cellFormat);
        }
        else {
            worksheet->write(excelRow, excelCol, cell->value, cellFormat);
        }

        processedCells++;
        if (totalCells > 0) {
            progress->setValue(10 + (processedCells * 70 / totalCells));
            qApp->processEvents();
        }
    }

    // 保存合并单元格
    saveMergedCells(worksheet, allCells);
    progress->setValue(90);

    if (progress->wasCanceled()) return false;

    if (!xlsx.saveAs(actualFileName)) {
        QMessageBox::warning(nullptr, "保存失败", QString("无法保存文件到：%1").arg(actualFileName));
        return false;
    }
    progress->setValue(100);
    return true;
}

void ExcelHandler::loadMergedCells(QXlsx::Worksheet* worksheet, QHash<QPoint, RTMergedRange>& mergedRanges)
{
    // 获取合并单元格范围列表
    const auto& mergedCells = worksheet->mergedCells();

    for (const QXlsx::CellRange& range : mergedCells) {
        if (range.isValid()) {
            RTMergedRange mergedRange(
                range.firstRow() - 1,    // 转换为0基索引
                range.firstColumn() - 1,
                range.lastRow() - 1,
                range.lastColumn() - 1
            );

            // 为合并范围内的所有单元格设置合并信息
            for (int row = range.firstRow(); row <= range.lastRow(); ++row) {
                for (int col = range.firstColumn(); col <= range.lastColumn(); ++col) {
                    QPoint cellPos(row - 1, col - 1);  // 转换为0基索引
                    mergedRanges[cellPos] = mergedRange;
                }
            }
        }
    }
}

void ExcelHandler::saveMergedCells(QXlsx::Worksheet* worksheet, const QHash<QPoint, RTCell*>& allCells)
{
    QSet<RTMergedRange*> processedRanges;  // 避免重复处理相同的合并范围

    for (auto it = allCells.constBegin(); it != allCells.constEnd(); ++it) {
        const RTCell* cell = it.value();

        if (cell->isMergedMain() && !processedRanges.contains(const_cast<RTMergedRange*>(&cell->mergedRange))) {
            // 转换为Excel的1基索引
            QXlsx::CellRange range(
                cell->mergedRange.startRow + 1,
                cell->mergedRange.startCol + 1,
                cell->mergedRange.endRow + 1,
                cell->mergedRange.endCol + 1
            );

            worksheet->mergeCells(range);
            processedRanges.insert(const_cast<RTMergedRange*>(&cell->mergedRange));
        }
    }
}


void ExcelHandler::convertFromExcelStyle(const QXlsx::Format& excelFormat, RTCellStyle& cellStyle)
{
    // 调试背景色读取
    qDebug() << "=== Background Color Debug ===";
    qDebug() << "Pattern background:" << excelFormat.patternBackgroundColor();
    qDebug() << "Pattern foreground:" << excelFormat.patternForegroundColor();
    qDebug() << "Fill pattern:" << excelFormat.fillPattern();

    // 尝试多种方式获取背景色
    QColor bgColor = excelFormat.patternBackgroundColor();
    if (!bgColor.isValid() || bgColor == QColor()) {
        bgColor = excelFormat.patternForegroundColor();
        qDebug() << "Using foreground as background";
    }
    if (!bgColor.isValid() || bgColor == QColor()) {
        bgColor = Qt::white;
        qDebug() << "Using default white";
    }

    // 检查是否是QXlsx读取中文字体的错误情况
    if (cellStyle.font.family() == "Calibri" && excelFormat.fontSize() == 0) {
        // 这很可能是中文字体被错误读取的情况
        cellStyle.font.setFamily("SimSun");  // 宋体的英文名
        cellStyle.font.setPointSize(11);      // 设置合理的默认字号
    }
    else {
        // 正常的英文字体，保持原有逻辑
        int fontSize = excelFormat.fontSize();
        if (fontSize > 0) {
            cellStyle.font.setPointSize(fontSize);
        }
        else if (cellStyle.font.pointSize() <= 0) {
            cellStyle.font.setPointSize(10); // 默认字号
        }
    }

    cellStyle.backgroundColor = bgColor;
    cellStyle.textColor = excelFormat.fontColor();


    // 对齐方式
    Qt::Alignment hAlign = Qt::AlignLeft;
    switch (excelFormat.horizontalAlignment()) {
    case QXlsx::Format::AlignHCenter: hAlign = Qt::AlignHCenter; break;
    case QXlsx::Format::AlignRight:   hAlign = Qt::AlignRight; break;
    default: break;
    }
    Qt::Alignment vAlign = Qt::AlignVCenter;
    switch (excelFormat.verticalAlignment()) {
    case QXlsx::Format::AlignTop:    vAlign = Qt::AlignTop; break;
    case QXlsx::Format::AlignBottom: vAlign = Qt::AlignBottom; break;
    default: break;
    }
    cellStyle.alignment = hAlign | vAlign;

    // 边框 - 现在可以正常处理了
    convertBorderFromExcel(excelFormat, cellStyle.border);
}

QXlsx::Format ExcelHandler::convertToExcelFormat(const RTCellStyle& cellStyle)
{
    QXlsx::Format excelFormat;

    // 字体
    excelFormat.setFont(cellStyle.font);

    // 颜色
    excelFormat.setPatternBackgroundColor(cellStyle.backgroundColor);
    excelFormat.setFontColor(cellStyle.textColor);

    // 对齐方式
    if (cellStyle.alignment & Qt::AlignHCenter)
        excelFormat.setHorizontalAlignment(QXlsx::Format::AlignHCenter);
    else if (cellStyle.alignment & Qt::AlignRight)
        excelFormat.setHorizontalAlignment(QXlsx::Format::AlignRight);

    if (cellStyle.alignment & Qt::AlignTop)
        excelFormat.setVerticalAlignment(QXlsx::Format::AlignTop);
    else if (cellStyle.alignment & Qt::AlignBottom)
        excelFormat.setVerticalAlignment(QXlsx::Format::AlignBottom);
    else
        excelFormat.setVerticalAlignment(QXlsx::Format::AlignVCenter);

    // 边框 - 现在可以正常处理了
    convertBorderToExcel(cellStyle.border, excelFormat);

    return excelFormat;
}

void ExcelHandler::convertBorderFromExcel(const QXlsx::Format& excelFormat, RTCellBorder& border)
{
    // 转换边框样式 - 使用正确的API
    border.left = convertBorderStyleFromExcel(excelFormat.leftBorderStyle());
    border.right = convertBorderStyleFromExcel(excelFormat.rightBorderStyle());
    border.top = convertBorderStyleFromExcel(excelFormat.topBorderStyle());
    border.bottom = convertBorderStyleFromExcel(excelFormat.bottomBorderStyle());

    // 转换边框颜色 - 使用正确的API
    border.leftColor = excelFormat.leftBorderColor();
    border.rightColor = excelFormat.rightBorderColor();
    border.topColor = excelFormat.topBorderColor();
    border.bottomColor = excelFormat.bottomBorderColor();
}


void ExcelHandler::convertBorderToExcel(const RTCellBorder& border, QXlsx::Format& excelFormat)
{
    // 转换边框样式 - 使用正确的API
    excelFormat.setLeftBorderStyle(convertBorderStyleToExcel(border.left));
    excelFormat.setRightBorderStyle(convertBorderStyleToExcel(border.right));
    excelFormat.setTopBorderStyle(convertBorderStyleToExcel(border.top));
    excelFormat.setBottomBorderStyle(convertBorderStyleToExcel(border.bottom));

    // 转换边框颜色 - 使用正确的API
    excelFormat.setLeftBorderColor(border.leftColor);
    excelFormat.setRightBorderColor(border.rightColor);
    excelFormat.setTopBorderColor(border.topColor);
    excelFormat.setBottomBorderColor(border.bottomColor);
}


RTBorderStyle ExcelHandler::convertBorderStyleFromExcel(QXlsx::Format::BorderStyle xlsxStyle)
{
    switch (xlsxStyle) {
    case QXlsx::Format::BorderNone: return RTBorderStyle::None;
    case QXlsx::Format::BorderThin: return RTBorderStyle::Thin;
    case QXlsx::Format::BorderMedium: return RTBorderStyle::Medium;
    case QXlsx::Format::BorderThick: return RTBorderStyle::Thick;
    case QXlsx::Format::BorderDouble: return RTBorderStyle::Double;
    case QXlsx::Format::BorderDotted: return RTBorderStyle::Dotted;
    case QXlsx::Format::BorderDashed: return RTBorderStyle::Dashed;
        // QXlsx还有更多边框样式，可以映射到我们的样式
    case QXlsx::Format::BorderHair: return RTBorderStyle::Thin;  // 映射为细线
    case QXlsx::Format::BorderMediumDashed: return RTBorderStyle::Dashed;
    case QXlsx::Format::BorderDashDot: return RTBorderStyle::Dashed;  // 映射为虚线
    case QXlsx::Format::BorderMediumDashDot: return RTBorderStyle::Dashed;
    case QXlsx::Format::BorderDashDotDot: return RTBorderStyle::Dashed;
    case QXlsx::Format::BorderMediumDashDotDot: return RTBorderStyle::Dashed;
    case QXlsx::Format::BorderSlantDashDot: return RTBorderStyle::Dashed;
    default: return RTBorderStyle::None;
    }
}

QXlsx::Format::BorderStyle ExcelHandler::convertBorderStyleToExcel(RTBorderStyle rtStyle)
{
    switch (rtStyle) {
    case RTBorderStyle::None: return QXlsx::Format::BorderNone;
    case RTBorderStyle::Thin: return QXlsx::Format::BorderThin;
    case RTBorderStyle::Medium: return QXlsx::Format::BorderMedium;
    case RTBorderStyle::Thick: return QXlsx::Format::BorderThick;
    case RTBorderStyle::Double: return QXlsx::Format::BorderDouble;
    case RTBorderStyle::Dotted: return QXlsx::Format::BorderDotted;
    case RTBorderStyle::Dashed: return QXlsx::Format::BorderDashed;
    default: return QXlsx::Format::BorderNone;
    }
}


bool ExcelHandler::isValidExcelFile(const QString& fileName)
{
    QFileInfo fileInfo(fileName);
    if (!fileInfo.exists()) {
        QMessageBox::warning(nullptr, "错误", QString("文件不存在：%1").arg(fileName));
        return false;
    }
    if (!fileInfo.isReadable()) {
        QMessageBox::warning(nullptr, "错误", QString("无法读取文件：%1").arg(fileName));
        return false;
    }
    return true;
}

