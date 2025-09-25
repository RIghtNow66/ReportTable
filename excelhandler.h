#ifndef EXCELHANDLER_H
#define EXCELHANDLER_H

#include <QString>

struct RTCell;
struct RTCellStyle; 
struct RTMergedRange;
class ReportDataModel;
class QXlsx;

namespace QXlsx {
    class Format;
    class Worksheet;
}

class ExcelHandler
{
public:
    static bool loadFromFile(const QString& fileName, ReportDataModel* model);
    static bool saveToFile(const QString& fileName, ReportDataModel* model);
private:
    ExcelHandler() = delete;
    ~ExcelHandler() = delete;
    static void convertFromExcelStyle(const QXlsx::Format& excelFormat, RTCellStyle& cellStyle);
    static QXlsx::Format convertToExcelFormat(const RTCellStyle& cellStyle);
    static bool isValidExcelFile(const QString& fileName);

    static void loadMergedCells(QXlsx::Worksheet* worksheet, QHash<QPoint, RTMergedRange>& mergedRanges);
    static void saveMergedCells(QXlsx::Worksheet* worksheet, const QHash<QPoint, RTCell*>& allCells);
    static void convertBorderFromExcel(const QXlsx::Format& excelFormat, RTCellBorder& border);
    static void convertBorderToExcel(const RTCellBorder& border, QXlsx::Format& excelFormat);
    RTBorderStyle convertBorderStyleFromExcel(QXlsx::Format::BorderStyle xlsxStyle);
    QXlsx::Format::BorderStyle convertBorderStyleToExcel(RTBorderStyle rtStyle);
};
#endif // EXCELHANDLER_H
