#include "formulaengine.h"
#include "reportdatamodel.h"
#include <QRegularExpression>
#include <QStack>
#include <QQueue>

QVariant FormulaEngine::evaluate(const QString& formula, ReportDataModel* model, int currentRow, int currentCol)
{
    Q_UNUSED(currentRow)
        Q_UNUSED(currentCol)

        if (!isFormula(formula))
            return formula;

    QString expr = formula.mid(1).trimmed(); // 去掉 '='

    // 优先检查是否为函数调用
    QRegularExpression functionRegex(R"(^(SUM|MAX|MIN)\s*\([^)]+\)$)");
    QRegularExpressionMatch functionMatch = functionRegex.match(expr);

    if (functionMatch.hasMatch()) {
        // 处理函数
        if (expr.startsWith("SUM")) {
            QRegularExpression sumRegex(R"(SUM\(([A-Z]+\d+:[A-Z]+\d+)\))");
            QRegularExpressionMatch match = sumRegex.match(expr);
            if (match.hasMatch()) {
                return evaluateSum(match.captured(1), model);
            }
        }
        else if (expr.startsWith("MAX")) {
            QRegularExpression maxRegex(R"(MAX\(([A-Z]+\d+:[A-Z]+\d+)\))");
            QRegularExpressionMatch match = maxRegex.match(expr);
            if (match.hasMatch()) {
                return evaluateMax(match.captured(1), model);
            }
        }
        else if (expr.startsWith("MIN")) {
            QRegularExpression minRegex(R"(MIN\(([A-Z]+\d+:[A-Z]+\d+)\))");
            QRegularExpressionMatch match = minRegex.match(expr);
            if (match.hasMatch()) {
                return evaluateMin(match.captured(1), model);
            }
        }
    }

    // 如果不是函数，尝试作为表达式处理
    return evaluateExpression(expr, model);
}

QVariant FormulaEngine::evaluateExpression(const QString& expression, ReportDataModel* model)
{
    // 1. 预处理：将单元格引用替换为实际值
    QString processedExpr = preprocessExpression(expression, model);

    // 2. 计算算术表达式
    return calculateArithmetic(processedExpr);
}

QString FormulaEngine::preprocessExpression(const QString& expression, ReportDataModel* model)
{
    QString result = expression;

    // 匹配单元格引用（如A1, B2, AA10等）
    QRegularExpression cellRefRegex(R"([A-Z]+\d+)");
    QRegularExpressionMatchIterator it = cellRefRegex.globalMatch(expression);

    // 从后往前替换，避免位置偏移
    QList<QRegularExpressionMatch> matches;
    while (it.hasNext()) {
        matches.append(it.next());
    }

    for (int i = matches.size() - 1; i >= 0; --i) {
        const QRegularExpressionMatch& match = matches[i];
        QString cellRef = match.captured(0);
        QVariant cellValue = getCellValue(cellRef, model);

        // 将单元格值转换为数字字符串
        bool ok;
        double numValue = cellValue.toDouble(&ok);
        if (!ok) {
            numValue = 0.0; // 如果不是数字，当作0处理
        }

        result.replace(match.capturedStart(), match.capturedLength(), QString::number(numValue));
    }

    return result;
}

QVariant FormulaEngine::calculateArithmetic(const QString& expression)
{
    // 使用调度场算法（Shunting Yard Algorithm）计算表达式
    QStack<double> numbers;
    QStack<QChar> operators;

    QString cleanExpr = expression;
    cleanExpr.remove(' '); // 移除空格

    int i = 0;
    while (i < cleanExpr.length()) {
        QChar c = cleanExpr[i];

        if (c.isDigit() || c == '.') {
            // 读取完整的数字
            QString numStr;
            while (i < cleanExpr.length() && (cleanExpr[i].isDigit() || cleanExpr[i] == '.')) {
                numStr += cleanExpr[i];
                i++;
            }
            numbers.push(numStr.toDouble());
            continue;
        }
        else if (c == '(') {
            operators.push(c);
        }
        else if (c == ')') {
            while (!operators.isEmpty() && operators.top() != '(') {
                QChar op = operators.pop();
                if (numbers.size() < 2) return QVariant(); // 错误
                double right = numbers.pop();
                double left = numbers.pop();
                numbers.push(applyOperator(left, right, op));
            }
            if (!operators.isEmpty()) operators.pop(); // 移除 '('
        }
        else if (isOperator(c)) {
            while (!operators.isEmpty() &&
                operators.top() != '(' &&
                getOperatorPrecedence(operators.top()) >= getOperatorPrecedence(c)) {
                QChar op = operators.pop();
                if (numbers.size() < 2) return QVariant(); // 错误
                double right = numbers.pop();
                double left = numbers.pop();
                numbers.push(applyOperator(left, right, op));
            }
            operators.push(c);
        }

        i++;
    }

    // 处理剩余的操作符
    while (!operators.isEmpty()) {
        QChar op = operators.pop();
        if (numbers.size() < 2) return QVariant(); // 错误
        double right = numbers.pop();
        double left = numbers.pop();
        numbers.push(applyOperator(left, right, op));
    }

    if (numbers.size() == 1) {
        return numbers.top();
    }

    return QVariant(); // 表达式错误
}

QVariant FormulaEngine::getCellValue(const QString& cellRef, ReportDataModel* model)
{
    QPoint pos = parseReference(cellRef);
    if (pos.x() == -1 || pos.y() == -1) {
        return QVariant(0.0);
    }

    const RTCell* cell = model->getCell(pos.x(), pos.y());
    if (!cell) {
        return QVariant(0.0);
    }

    return cell->value;
}

double FormulaEngine::applyOperator(double left, double right, QChar op)
{
    switch (op.unicode()) {
    case '+': return left + right;
    case '-': return left - right;
    case '*': return left * right;
    case '/': return (right != 0) ? left / right : 0;
    default: return 0;
    }
}

int FormulaEngine::getOperatorPrecedence(QChar op)
{
    switch (op.unicode()) {
    case '+': case '-': return 1;
    case '*': case '/': return 2;
    default: return 0;
    }
}

bool FormulaEngine::isOperator(QChar c)
{
    return c == '+' || c == '-' || c == '*' || c == '/';
}

bool FormulaEngine::isValidCellReference(const QString& ref) const
{
    QRegularExpression regex(R"(^[A-Z]+\d+$)");
    return regex.match(ref).hasMatch();
}