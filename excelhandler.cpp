#include "excelhandler.h"
#include "reportdatamodel.h"
#include "Cell.h"

#include <QMessageBox>
#include <QFileInfo>
#include <QApplication>
#include <QProgressDialog>
#include <memory> // for std::make_unique

// 明确包含所有需要的QXlsx库头文件
#include "xlsxdocument.h"
#include "xlsxworksheet.h"
#include "xlsxcell.h"
#include "xlsxformat.h"
#include "xlsxcellrange.h"
// 修正：移除不存在的 "xlsxformula.h"

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
        model->updateModelSize(0, 0); // 更新为空白表格
        return true;
    }

    int totalRows = range.rowCount();
    int maxModelRow = 0, maxModelCol = 0;

    for (int row = range.firstRow(); row <= range.lastRow(); ++row) {
        if (progress->wasCanceled()) break;
        for (int col = range.firstColumn(); col <= range.lastColumn(); ++col) {

            auto xlsxCell = worksheet->cellAt(row, col);

            if (xlsxCell) {
                RTCell* newCell = new RTCell();

                // 最终修正: 采用最可靠的方式读取公式，以完全绕开库版本问题
                // 1. 读取单元格的原始数据
                QVariant rawData = worksheet->read(row, col);

                // 2. 检查原始数据是否为字符串，并且以'='开头
                if (rawData.type() == QVariant::String && rawData.toString().startsWith('=')) {
                    newCell->setFormula(rawData.toString());
                }

                newCell->value = xlsxCell->value(); // 无论如何都获取公式计算后的值或普通值

                if (xlsxCell->format().isValid()) {
                    convertFromExcelStyle(xlsxCell->format(), newCell->style);
                }
                int modelRow = row - 1, modelCol = col - 1;
                model->addCellDirect(modelRow, modelCol, newCell);
                maxModelRow = qMax(maxModelRow, modelRow + 1);
                maxModelCol = qMax(maxModelCol, modelCol + 1);
            }
        }
        int progressValue = 20 + ((row - range.firstRow() + 1) * 70 / totalRows);
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
            progress->setValue(10 + (processedCells * 80 / totalCells));
            qApp->processEvents();
        }
    }
    if (progress->wasCanceled()) return false;
    progress->setValue(95);

    if (!xlsx.saveAs(actualFileName)) {
        QMessageBox::warning(nullptr, "保存失败", QString("无法保存文件到：%1").arg(actualFileName));
        return false;
    }
    progress->setValue(100);
    return true;
}

void ExcelHandler::convertFromExcelStyle(const QXlsx::Format& excelFormat, RTCellStyle& cellStyle)
{
    cellStyle.font = excelFormat.font();
    cellStyle.backgroundColor = excelFormat.patternBackgroundColor();
    cellStyle.textColor = excelFormat.fontColor();
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
}

QXlsx::Format ExcelHandler::convertToExcelFormat(const RTCellStyle& cellStyle)
{
    QXlsx::Format excelFormat;
    excelFormat.setFont(cellStyle.font);
    excelFormat.setPatternBackgroundColor(cellStyle.backgroundColor);
    excelFormat.setFontColor(cellStyle.textColor);
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
    return excelFormat;
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

