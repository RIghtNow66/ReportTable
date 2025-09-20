#ifndef FORMULAENGINE_H
#define FORMULAENGINE_H

#include <QObject>
#include <QVariant>
#include <QString>
#include "Cell.h"
class ReportDataModel;

class FormulaEngine : public QObject
{
    Q_OBJECT

public:
    explicit FormulaEngine(QObject* parent = nullptr);

    QVariant evaluate(const QString& formula, ReportDataModel* model, int currentRow, int currentCol);
    bool isFormula(const QString& text) const;

private:
    QVariant evaluateSum(const QString& range, ReportDataModel* model);
    QVariant evaluateMax(const QString& range, ReportDataModel* model);
    QVariant evaluateMin(const QString& range, ReportDataModel* model);

    QPoint parseReference(const QString& ref) const;
    QPair<QPoint, QPoint> parseRange(const QString& range) const;
};

#endif // FORMULAENGINE_H
