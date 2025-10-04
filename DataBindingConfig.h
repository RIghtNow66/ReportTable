#pragma once
#ifndef DATABINDINGCONFIG_H
#define DATABINDINGCONFIG_H

#include <QDateTime>
#include <QString>
#include <QVariant>
#include <QFont>
#include <QSet>
#include <QColor>

// ===== ��ʽ�ṹ����Cell.h�ƹ����� =====
enum class RTBorderStyle {
    None = 0, Thin, Medium, Thick, Double, Dotted, Dashed
};

struct RTCellBorder {
    RTBorderStyle left = RTBorderStyle::None;
    RTBorderStyle right = RTBorderStyle::None;
    RTBorderStyle top = RTBorderStyle::None;
    RTBorderStyle bottom = RTBorderStyle::None;
    QColor leftColor = Qt::black;
    QColor rightColor = Qt::black;
    QColor topColor = Qt::black;
    QColor bottomColor = Qt::black;

    RTCellBorder() = default;
};

struct RTCellStyle {
    QFont font;
    QColor backgroundColor;
    QColor textColor = Qt::black;
    Qt::Alignment alignment = Qt::AlignLeft | Qt::AlignVCenter;
    RTCellBorder border;

    RTCellStyle() {
        font.setFamily("Arial");
        font.setPointSize(10);
        backgroundColor = Qt::white;
        font.setBold(false);
    }
};

// ===== �ϲ���Ϣ����Cell.h�ƹ����� =====
struct RTMergedRange {
    int startRow;
    int startCol;
    int endRow;
    int endCol;

    RTMergedRange() :startRow(-1), startCol(-1), endRow(-1), endCol(-1) {}
    RTMergedRange(int sRow, int sCol, int eRow, int eCol)
        :startRow(sRow), startCol(sCol), endRow(eRow), endCol(eCol) {
    }

    bool isValid() const {
        return startRow >= 0 && startCol >= 0 && endRow >= startRow && endCol >= startCol;
    }

    bool isMerged() const {
        return isValid() && (startRow != endRow || startCol != endCol);
    }

    bool contains(int row, int col) const {
        return isValid() && row >= startRow && row <= endRow && col >= startCol && col <= endCol;
    }

    int rowSpan() const { return isValid() ? (endRow - startRow + 1) : 1; }
    int colSpan() const { return isValid() ? (endCol - startCol + 1) : 1; }
};

// ===== ͳһ�ĵ�Ԫ�����ݽṹ =====
struct CellData {
    // 1. ��������
    QVariant value;
    QString formula;
    bool hasFormula = false;

    // 2. ��ʽ
    RTCellStyle style;

    // 3. �ϲ���Ϣ
    RTMergedRange mergedRange;

    // 4. ���ݰ󶨣������󶨣�
    bool isDataBinding = false;
    QString bindingKey;          // �����󶨼����� "##TEMP_01"

    CellData() = default;

    // ���߷���
    void setFormula(const QString& formulaText) {
        formula = formulaText;
        hasFormula = !formulaText.isEmpty() && formulaText.startsWith('=');
    }

    QString displayText() const {
        return value.toString();
    }

    QString editText() const {
        if (isDataBinding) {
            return bindingKey;  // ֱ�ӷ���bindingKey
        }
        return hasFormula ? formula : value.toString();
    }

    bool isMergedMain() const {
        return mergedRange.isValid() && mergedRange.isMerged();
    }

    bool isMergedChild() const {
        return mergedRange.isValid() && isMergedMain() &&
            (mergedRange.startRow != 0 || mergedRange.startCol != 0);
    }
};

// ===== ʱ�䷶Χ���� =====
struct TimeRangeConfig {
    QDateTime startTime;
    QDateTime endTime;
    int intervalSeconds;

    TimeRangeConfig()
        : intervalSeconds(300)
    {
        QDateTime now = QDateTime::currentDateTime();
        startTime = QDateTime(now.date(), QTime(0, 0, 0));
        endTime = now;
    }

    TimeRangeConfig(const QDateTime& start, const QDateTime& end, int interval)
        : startTime(start), endTime(end), intervalSeconds(interval)
    {
    }

    bool isValid() const {
        return startTime.isValid() && endTime.isValid()
            && startTime < endTime && intervalSeconds > 0;
    }
};

// ===== ȫ������ =====
struct GlobalDataConfig {
    TimeRangeConfig globalTimeRange;  // ȫ��ʱ�䷶Χ
    QString defaultRTU;               // Ĭ��RTU��

    GlobalDataConfig()
        : defaultRTU("AIRTU000100006")
    {
    }
};

// ����������
struct ReportColumnConfig {
    QString displayName;     // ��ʾ����
    QString rtuId;           // RTU��
    int sourceRow;           // �����ļ��е��к�

    ReportColumnConfig() : sourceRow(-1) {}
};

// ��ʷ��������
struct HistoryReportConfig {
    QVector<ReportColumnConfig> columns;
    QString reportName;      // ��������
    QString configFilePath;  // �����ļ�·��

    // ��¼��Щ���������У�ֻ����
    QSet<int> dataColumns;  // �洢�����е�������dataColumns 

    HistoryReportConfig() {}
};


#endif // DATABINDINGCONFIG_H