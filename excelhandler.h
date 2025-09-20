#ifndef EXCELHANDLER_H
#define EXCELHANDLER_H

#include <QString>

struct RTCell;
struct RTCellStyle; // 淇: 涓嶳TCell淇濇寔涓€鑷?
class ReportDataModel;

namespace QXlsx {
    class Format;
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
};
#endif // EXCELHANDLER_H
