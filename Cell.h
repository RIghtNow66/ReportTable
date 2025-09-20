#ifndef RTCELL_H
#define RTCELL_H

#include <QVariant>
#include <QString>
#include <QFont>
#include <QColor>

struct RTCellStyle {
    QFont font;
    QColor backgroundColor = Qt::white;
    QColor textColor = Qt::black;
    Qt::Alignment alignment = Qt::AlignLeft | Qt::AlignVCenter;

    RTCellStyle() {
        font.setFamily("Arial");
        font.setPointSize(10);
    }
};

struct RTCell {
    QVariant value;
    QString formula;
    RTCellStyle style;
    bool hasFormula = false;

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
};

#endif // RTCELL_H
