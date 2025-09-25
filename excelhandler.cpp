#include "excelhandler.h"
#include "reportdatamodel.h"
#include "Cell.h"

#include <QMessageBox>
#include <QFileInfo>
#include <QApplication>
#include <QProgressDialog>
#include <memory> // for std::make_unique
#include <QSet>

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
    progress->setValue(100);
    return true;
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
    // 字体
    cellStyle.font = excelFormat.font();

    // 颜色
    cellStyle.backgroundColor = excelFormat.patternBackgroundColor();
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

    // 边框
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

    // 边框
    convertBorderToExcel(cellStyle.border, excelFormat);

    return excelFormat;
}

void ExcelHandler::convertBorderFromExcel(const QXlsx::Format& excelFormat, RTCellBorder& border)
{
    // 转换边框样式
    border.left = convertBorderStyleFromExcel(excelFormat.borderStyle(QXlsx::Format::BorderLeft));
    border.right = convertBorderStyleFromExcel(excelFormat.borderStyle(QXlsx::Format::BorderRight));
    border.top = convertBorderStyleFromExcel(excelFormat.borderStyle(QXlsx::Format::BorderTop));
    border.bottom = convertBorderStyleFromExcel(excelFormat.borderStyle(QXlsx::Format::BorderBottom));

    // 转换边框颜色
    border.leftColor = excelFormat.borderColor(QXlsx::Format::BorderLeft);
    border.rightColor = excelFormat.borderColor(QXlsx::Format::BorderRight);
    border.topColor = excelFormat.borderColor(QXlsx::Format::BorderTop);
    border.bottomColor = excelFormat.borderColor(QXlsx::Format::BorderBottom);
}

void ExcelHandler::convertBorderToExcel(const RTCellBorder& border, QXlsx::Format& excelFormat)
{
    // 转换边框样式
    excelFormat.setBorderStyle(QXlsx::Format::BorderLeft, convertBorderStyleToExcel(border.left));
    excelFormat.setBorderStyle(QXlsx::Format::BorderRight, convertBorderStyleToExcel(border.right));
    excelFormat.setBorderStyle(QXlsx::Format::BorderTop, convertBorderStyleToExcel(border.top));
    excelFormat.setBorderStyle(QXlsx::Format::BorderBottom, convertBorderStyleToExcel(border.bottom));

    // 转换边框颜色
    excelFormat.setBorderColor(QXlsx::Format::BorderLeft, border.leftColor);
    excelFormat.setBorderColor(QXlsx::Format::BorderRight, border.rightColor);
    excelFormat.setBorderColor(QXlsx::Format::BorderTop, border.topColor);
    excelFormat.setBorderColor(QXlsx::Format::BorderBottom, border.bottomColor);
}

RTBorderStyle ExcelHandler::convertBorderStyleFromExcel(QXlsx::Format::BorderStyle xlsxStyle)
{
    switch (xlsxStyle) {
    case QXlsx::Format::BorderThin: return RTBorderStyle::Thin;
    case QXlsx::Format::BorderMedium: return RTBorderStyle::Medium;
    case QXlsx::Format::BorderThick: return RTBorderStyle::Thick;
    case QXlsx::Format::BorderDouble: return RTBorderStyle::Double;
    case QXlsx::Format::BorderDotted: return RTBorderStyle::Dotted;
    case QXlsx::Format::BorderDashed: return RTBorderStyle::Dashed;
    default: return RTBorderStyle::None;
    }
}

QXlsx::Format::BorderStyle ExcelHandler::convertBorderStyleToExcel(RTBorderStyle rtStyle)
{
    switch (rtStyle) {
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

