#include "formulaengine.h"
#include "reportdatamodel.h"
#include <QRegularExpression>

FormulaEngine::FormulaEngine(QObject* parent)
    : QObject(parent)
{
}

QVariant FormulaEngine::evaluate(const QString& formula, ReportDataModel* model, int currentRow, int currentCol)
{
    Q_UNUSED(currentRow)
        Q_UNUSED(currentCol)

        if (!isFormula(formula))
            return formula;

    QString expr = formula.mid(1).trimmed(); // 去掉 '='

    // 简单的SUM函数处理
    QRegularExpression sumRegex(R"(SUM\(([A-Z]+\d+:[A-Z]+\d+)\))");
    QRegularExpressionMatch match = sumRegex.match(expr);

    if (match.hasMatch()) {
        QString range = match.captured(1);
        return evaluateSum(range, model);
    }

    // 简单的MAX函数处理
    QRegularExpression maxRegex(R"(MAX\(([A-Z]+\d+:[A-Z]+\d+)\))");
    match = maxRegex.match(expr);

    if (match.hasMatch()) {
        QString range = match.captured(1);
        return evaluateMax(range, model);
    }

    // 简单的MIN函数处理
    QRegularExpression minRegex(R"(MIN\(([A-Z]+\d+:[A-Z]+\d+)\))");
    match = minRegex.match(expr);

    if (match.hasMatch()) {
        QString range = match.captured(1);
        return evaluateMin(range, model);
    }

    return formula;
}

bool FormulaEngine::isFormula(const QString& text) const
{
    return !text.isEmpty() && text.startsWith('=');
}

QVariant FormulaEngine::evaluateSum(const QString& range, ReportDataModel* model)
{
    QPair<QPoint, QPoint> rangePair = parseRange(range);
    QPoint start = rangePair.first;
    QPoint end = rangePair.second;

    double sum = 0.0;
    bool hasValidData = false;

    for (int row = start.x(); row <= end.x(); ++row) {
        for (int col = start.y(); col <= end.y(); ++col) {
            const RTCell* cell = model->getCell(row, col);
            if (cell) {
                bool ok;
                double value = cell->value.toDouble(&ok);
                if (ok) {
                    sum += value;
                    hasValidData = true;
                }
            }
        }
    }

    return hasValidData ? QVariant(sum) : QVariant(0);
}

QVariant FormulaEngine::evaluateMax(const QString& range, ReportDataModel* model)
{
    QPair<QPoint, QPoint> rangePair = parseRange(range);
    QPoint start = rangePair.first;
    QPoint end = rangePair.second;

    double maxVal = -std::numeric_limits<double>::infinity();
    bool hasValidData = false;

    for (int row = start.x(); row <= end.x(); ++row) {
        for (int col = start.y(); col <= end.y(); ++col) {
            const RTCell* cell = model->getCell(row, col);
            if (cell) {
                bool ok;
                double value = cell->value.toDouble(&ok);
                if (ok) {
                    maxVal = qMax(maxVal, value);
                    hasValidData = true;
                }
            }
        }
    }

    return hasValidData ? QVariant(maxVal) : QVariant();
}

QVariant FormulaEngine::evaluateMin(const QString& range, ReportDataModel* model)
{
    QPair<QPoint, QPoint> rangePair = parseRange(range);
    QPoint start = rangePair.first;
    QPoint end = rangePair.second;

    double minVal = std::numeric_limits<double>::infinity();
    bool hasValidData = false;

    for (int row = start.x(); row <= end.x(); ++row) {
        for (int col = start.y(); col <= end.y(); ++col) {
            const RTCell* cell = model->getCell(row, col);
            if (cell) {
                bool ok;
                double value = cell->value.toDouble(&ok);
                if (ok) {
                    minVal = qMin(minVal, value);
                    hasValidData = true;
                }
            }
        }
    }

    return hasValidData ? QVariant(minVal) : QVariant();
}

QPoint FormulaEngine::parseReference(const QString& ref) const
{
    QRegularExpression regex(R"(([A-Z]+)(\d+))");
    QRegularExpressionMatch match = regex.match(ref);

    if (!match.hasMatch())
        return QPoint(-1, -1);

    QString colStr = match.captured(1);
    int row = match.captured(2).toInt() - 1;

    int col = 0;
    for (QChar c : colStr) {
        col = col * 26 + (c.unicode() - 'A' + 1);
    }
    col -= 1;

    return QPoint(row, col);
}

QPair<QPoint, QPoint> FormulaEngine::parseRange(const QString& range) const
{
    QStringList parts = range.split(':');
    if (parts.size() != 2)
        return qMakePair(QPoint(-1, -1), QPoint(-1, -1));

    QPoint start = parseReference(parts[0]);
    QPoint end = parseReference(parts[1]);

    return qMakePair(start, end);
}
