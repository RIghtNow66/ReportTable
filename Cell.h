#ifndef RTCELL_H
#define RTCELL_H

#include <QVariant>
#include <QString>
#include <QFont>
#include <QColor>

enum class RTBorderStyle {
    None=0, // 无边框
    Thin,  // 细实线
    Medium, // 中等实线
    Thick, // 粗实线
    Double, // 双线
    Dotted, // 点线
    Dashed, // 虚线
};

// 单元格边框信息
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

// 单元格合并信息
struct RTMergedRange {
    int startRow;
    int startCol;
    int endRow;
    int endCol;

	RTMergedRange() :startRow(-1), startCol(-1), endRow(-1), endCol(-1) {}
	RTMergedRange(int sRow, int sCol, int eRow, int eCol) :startRow(sRow), startCol(sCol), endRow(eRow), endCol(eCol) {}

    bool isValid() const {
        return startRow >= 0 && startCol >= 0 && endRow >= startRow && endCol >= startCol;
	}

    bool isMerged() const {
        return isValid() && (startRow != endRow || startCol != endCol);
    }

	// 判断某个单元格是否在合并区域内
    bool contains(int row, int col) const {
		return isValid() && row >= startRow && row <= endRow && col >= startCol && col <= endCol;
    }

	// 获取合并区域的行数和列数
    int rowSpan() const { return isValid() ? (endRow - startRow + 1) : 1; }
    int colSpan() const { return isValid() ? (endCol - startCol + 1) : 1; }
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
        backgroundColor = Qt::white;  // 明确设置白色背景
    }
};

struct RTCell {
    QVariant value;
    QString formula;
    RTCellStyle style;
    bool hasFormula = false;
	RTMergedRange mergedRange;

    RTCell() = default;

    void setFormula(const QString& formulaText) {
        formula = formulaText;
        hasFormula = !formulaText.isEmpty() && formulaText.startsWith('=');
    }

    QString displayText() const {
        return value.toString();
    }

    QString editText() const {
        return hasFormula ? formula : value.toString();
    }

    bool isMergedMain() const
    {
        return mergedRange.isValid() && mergedRange.isMerged();
    }

    bool isMergedChild() const {
        return mergedRange.isValid() && isMergedMain() &&
            (mergedRange.startRow != 0 || mergedRange.startCol != 0);
    }
};

#endif // RTCELL_H
