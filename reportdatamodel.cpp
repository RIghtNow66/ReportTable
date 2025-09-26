#include "reportdatamodel.h"
#include "formulaengine.h"
#include "excelhandler.h" // 用于文件操作

#include <QColor>
#include <QFont>
#include <QPoint>
#include <QHash>
#include <QBrush>
#include <QPen>

// --- 构造与析构 ---

ReportDataModel::ReportDataModel(QObject* parent)
    : QAbstractTableModel(parent)
    , m_maxRow(100) // 默认初始行数
    , m_maxCol(26)  // 默认初始列数 (A-Z)
    , m_formulaEngine(new FormulaEngine(this))
{
}

ReportDataModel::~ReportDataModel()
{
    qDeleteAll(m_cells);
}

// --- Qt Model 核心接口实现 ---

int ReportDataModel::rowCount(const QModelIndex& parent) const
{
    Q_UNUSED(parent)
        return m_maxRow;
}

int ReportDataModel::columnCount(const QModelIndex& parent) const
{
    Q_UNUSED(parent)
        return m_maxCol;
}

QSize ReportDataModel::span(const QModelIndex& index) const
{
    if (!index.isValid())
        return QSize(1, 1);

    const RTCell* cell = getCell(index.row(), index.column());
    if (!cell || !cell->mergedRange.isValid() || !cell->mergedRange.isMerged())
        return QSize(1, 1);

    // 只有主单元格返回span信息
    if (index.row() == cell->mergedRange.startRow &&
        index.column() == cell->mergedRange.startCol) {
        return QSize(cell->mergedRange.colSpan(), cell->mergedRange.rowSpan());
    }

    return QSize(1, 1);
}

QFont ReportDataModel::ensureFontAvailable(const QFont& requestedFont) const
{
    QFont font = requestedFont;

    // 检查字体是否在系统中可用
    QFontInfo fontInfo(font);
    QString actualFamily = fontInfo.family();
    QString requestedFamily = font.family();

    if (actualFamily != requestedFamily) {

        // 中文字体映射
        if (requestedFamily.contains("宋体") || requestedFamily == "SimSun") {
            QStringList alternatives = { "SimSun", "NSimSun", "宋体", "新宋体" };
            for (const QString& alt : alternatives) {
                font.setFamily(alt);
                QFontInfo altInfo(font);
                if (altInfo.family().contains(alt, Qt::CaseInsensitive)) {
                    return font;
                }
            }
        }

        if (requestedFamily.contains("黑体") || requestedFamily == "SimHei") {
            QStringList alternatives = { "Microsoft YaHei", "SimHei", "黑体" };
            for (const QString& alt : alternatives) {
                font.setFamily(alt);
                QFontInfo altInfo(font);
                if (altInfo.family().contains(alt, Qt::CaseInsensitive)) {
                    return font;
                }
            }
        }
        font.setFamily("");
    }

    return font;
}

QVariant ReportDataModel::data(const QModelIndex& index, int role) const
{
    if (!index.isValid())
        return QVariant();

    const RTCell* cell = getCell(index.row(), index.column());
    if (!cell)
        return QVariant();

    // 对于合并单元格处理
    if (cell->mergedRange.isValid() && cell->mergedRange.isMerged()) {
        bool isMainCell = (index.row() == cell->mergedRange.startRow &&
            index.column() == cell->mergedRange.startCol);

        switch (role) {
        case Qt::DisplayRole:
        case Qt::EditRole:
            return isMainCell ? cell->value.toString() : QVariant();
        case Qt::BackgroundRole:
            return QBrush(cell->style.backgroundColor);
        case Qt::ForegroundRole:
            return QBrush(cell->style.textColor);
        case Qt::FontRole: {
            QFont font = ensureFontAvailable(cell->style.font);
            return font;
        }
        case Qt::TextAlignmentRole:
            return static_cast<int>(cell->style.alignment);
        default:
            return QVariant();
        }
    }

    // 普通单元格处理
    switch (role) {
    case Qt::DisplayRole:
        return cell->displayText();
    case Qt::EditRole:
        return cell->editText();
    case Qt::BackgroundRole:
        return QBrush(cell->style.backgroundColor);
    case Qt::ForegroundRole:
        return QBrush(cell->style.textColor);
    case Qt::FontRole: {
        QFont font = ensureFontAvailable(cell->style.font);
        return font;
    }
    case Qt::TextAlignmentRole:
        return static_cast<int>(cell->style.alignment);
    default:
        return QVariant();
    }
}

bool ReportDataModel::setData(const QModelIndex& index, const QVariant& value, int role)
{
    if (!index.isValid() || role != Qt::EditRole)
        return false;

    RTCell* cell = ensureCell(index.row(), index.column());
    if (!cell) return false;

    QString text = value.toString();

    if (text.startsWith('=')) {
        cell->setFormula(text);
        calculateFormula(index.row(), index.column());
    }
    else {
        // 如果之前是公式，现在清空
        if (cell->hasFormula) {
            cell->formula.clear();
            cell->hasFormula = false;
        }
        cell->value = value;
    }

    // 通知视图更新
    emit dataChanged(index, index, { role });
    emit cellChanged(index.row(), index.column());
    return true;
}

Qt::ItemFlags ReportDataModel::flags(const QModelIndex& index) const
{
    if (!index.isValid())
        return Qt::NoItemFlags;

    // 检查是否是合并单元格的从属单元格
    const RTCell* cell = getCell(index.row(), index.column());
    if (cell && cell->mergedRange.isValid() && cell->mergedRange.isMerged()) {
        // 如果不是主单元格，则不可编辑
        if (!(index.row() == cell->mergedRange.startRow &&
            index.column() == cell->mergedRange.startCol)) {
            return Qt::ItemIsEnabled | Qt::ItemIsSelectable;
        }
    }

    return Qt::ItemIsEnabled | Qt::ItemIsSelectable | Qt::ItemIsEditable;
}
QVariant ReportDataModel::headerData(int section, Qt::Orientation orientation, int role) const
{
    if (role != Qt::DisplayRole)
        return QVariant();

    if (orientation == Qt::Horizontal) {
        QString colName;
        int temp = section;
        while (temp >= 0) {
            colName.prepend(QChar('A' + (temp % 26)));
            temp = temp / 26 - 1;
        }
        return colName;
    }
    else {
        return section + 1; // 行号从1开始
    }
}

// --- 行列操作实现 ---

bool ReportDataModel::insertRows(int row, int count, const QModelIndex& parent)
{
    Q_UNUSED(parent)
        if (count <= 0) return false;

    beginInsertRows(QModelIndex(), row, row + count - 1);

    QHash<QPoint, RTCell*> newCells;
    for (auto it = m_cells.begin(); it != m_cells.end(); ++it) {
        QPoint oldPos = it.key();
        RTCell* cell = it.value();

        if (oldPos.x() >= row) {
            // 将此行及以下的单元格向下移动
            QPoint newPos(oldPos.x() + count, oldPos.y());
            newCells[newPos] = cell;

            // 更新合并单元格信息
            if (cell->mergedRange.isValid()) {
                if (cell->mergedRange.startRow >= row) {
                    cell->mergedRange.startRow += count;
                    cell->mergedRange.endRow += count;
                }
            }
        }
        else {
            newCells[oldPos] = cell;

            // 更新跨越插入行的合并单元格信息
            if (cell->mergedRange.isValid() &&
                cell->mergedRange.startRow < row && cell->mergedRange.endRow >= row) {
                cell->mergedRange.endRow += count;
            }
        }
    }
    m_cells = newCells;
    m_maxRow += count;

    endInsertRows();
    return true;
}

bool ReportDataModel::removeRows(int row, int count, const QModelIndex& parent)
{
    Q_UNUSED(parent)
        if (count <= 0 || row < 0 || row + count > m_maxRow)
            return false;

    beginRemoveRows(QModelIndex(), row, row + count - 1);

    QHash<QPoint, RTCell*> newCells;
    for (auto it = m_cells.begin(); it != m_cells.end(); ++it) {
        QPoint oldPos = it.key();
        RTCell* cell = it.value();

        if (oldPos.x() >= row && oldPos.x() < row + count) {
            // 删除被移除范围内的单元格
            delete cell;
        }
        else if (oldPos.x() >= row + count) {
            // 将被移除范围下方的单元格向上移动
            QPoint newPos(oldPos.x() - count, oldPos.y());
            newCells[newPos] = cell;

            // 更新合并单元格信息
            if (cell->mergedRange.isValid()) {
                if (cell->mergedRange.startRow >= row + count) {
                    cell->mergedRange.startRow -= count;
                    cell->mergedRange.endRow -= count;
                }
            }
        }
        else {
            newCells[oldPos] = cell;

            // 更新跨越删除行的合并单元格信息
            if (cell->mergedRange.isValid()) {
                if (cell->mergedRange.endRow >= row + count) {
                    cell->mergedRange.endRow -= count;
                }
                else if (cell->mergedRange.endRow >= row) {
                    cell->mergedRange.endRow = row - 1;
                }

                // 如果合并范围无效了，清除合并信息
                if (cell->mergedRange.endRow < cell->mergedRange.startRow ||
                    cell->mergedRange.endCol < cell->mergedRange.startCol) {
                    cell->mergedRange = RTMergedRange();
                }
            }
        }
    }
    m_cells = newCells;
    m_maxRow -= count;

    endRemoveRows();
    return true;
}

bool ReportDataModel::insertColumns(int column, int count, const QModelIndex& parent)
{
    Q_UNUSED(parent)
        if (count <= 0) return false;

    beginInsertColumns(QModelIndex(), column, column + count - 1);

    QHash<QPoint, RTCell*> newCells;
    for (auto it = m_cells.begin(); it != m_cells.end(); ++it) {
        QPoint oldPos = it.key();
        RTCell* cell = it.value();

        if (oldPos.y() >= column) {
            QPoint newPos(oldPos.x(), oldPos.y() + count);
            newCells[newPos] = cell;

            // 更新合并单元格信息
            if (cell->mergedRange.isValid()) {
                if (cell->mergedRange.startCol >= column) {
                    cell->mergedRange.startCol += count;
                    cell->mergedRange.endCol += count;
                }
            }
        }
        else {
            newCells[oldPos] = cell;

            // 更新跨越插入列的合并单元格信息
            if (cell->mergedRange.isValid() &&
                cell->mergedRange.startCol < column && cell->mergedRange.endCol >= column) {
                cell->mergedRange.endCol += count;
            }
        }
    }
    m_cells = newCells;
    m_maxCol += count;

    endInsertColumns();
    return true;
}

bool ReportDataModel::removeColumns(int column, int count, const QModelIndex& parent)
{
    Q_UNUSED(parent)
        if (count <= 0 || column < 0 || column + count > m_maxCol)
            return false;

    beginRemoveColumns(QModelIndex(), column, column + count - 1);

    QHash<QPoint, RTCell*> newCells;
    for (auto it = m_cells.begin(); it != m_cells.end(); ++it) {
        QPoint oldPos = it.key();
        RTCell* cell = it.value();

        if (oldPos.y() >= column && oldPos.y() < column + count) {
            delete cell;
        }
        else if (oldPos.y() >= column + count) {
            QPoint newPos(oldPos.x(), oldPos.y() - count);
            newCells[newPos] = cell;

            // 更新合并单元格信息
            if (cell->mergedRange.isValid()) {
                if (cell->mergedRange.startCol >= column + count) {
                    cell->mergedRange.startCol -= count;
                    cell->mergedRange.endCol -= count;
                }
            }
        }
        else {
            newCells[oldPos] = cell;

            // 更新跨越删除列的合并单元格信息
            if (cell->mergedRange.isValid()) {
                if (cell->mergedRange.endCol >= column + count) {
                    cell->mergedRange.endCol -= count;
                }
                else if (cell->mergedRange.endCol >= column) {
                    cell->mergedRange.endCol = column - 1;
                }

                // 如果合并范围无效了，清除合并信息
                if (cell->mergedRange.endRow < cell->mergedRange.startRow ||
                    cell->mergedRange.endCol < cell->mergedRange.startCol) {
                    cell->mergedRange = RTMergedRange();
                }
            }
        }
    }
    m_cells = newCells;
    m_maxCol -= count;

    endRemoveColumns();
    return true;
}

// --- 文件操作实现 ---

bool ReportDataModel::loadFromExcel(const QString& fileName)
{
    // begin/endResetModel必须在模型内部调用，因为它们是受保护的
    beginResetModel();
    // 委托给ExcelHandler处理真正的文件读写逻辑
    bool result = ExcelHandler::loadFromFile(fileName, this);
    endResetModel();

    // 加载成功后，统一重新计算所有公式
    if (result) {
        recalculateAllFormulas();
    }
    return result;
}

bool ReportDataModel::saveToExcel(const QString& fileName)
{
    // 直接委托给ExcelHandler处理
    return ExcelHandler::saveToFile(fileName, this);
}

// --- 公共接口实现 (供ExcelHandler等外部类使用) ---

void ReportDataModel::clearAllCells()
{
    if (m_cells.isEmpty()) return;

    // 注意：这里不调用begin/endResetModel，因为调用方(loadFromExcel)会负责
    qDeleteAll(m_cells);
    m_cells.clear();
}

void ReportDataModel::addCellDirect(int row, int col, RTCell* cell)
{
    // 此方法专为Excel高速加载设计，不触发信号
    QPoint key(row, col);
    if (m_cells.contains(key)) {
        delete m_cells.take(key);
    }
    m_cells.insert(key, cell);
}

void ReportDataModel::updateModelSize(int newRowCount, int newColCount)
{
    // 更新内部行列数，保证最小尺寸
    m_maxRow = qMax(100, newRowCount);
    m_maxCol = qMax(26, newColCount);
    // 注意：同样不调用begin/endResetModel，由调用方负责
}

const QHash<QPoint, RTCell*>& ReportDataModel::getAllCells() const
{
    return m_cells;
}

void ReportDataModel::recalculateAllFormulas()
{
    for (auto it = m_cells.constBegin(); it != m_cells.constEnd(); ++it) {
        if (it.value() && it.value()->hasFormula) {
            const QPoint& pos = it.key();
            calculateFormula(pos.x(), pos.y());
        }
    }
    // 计算完成后，通知视图全局刷新数据
    emit dataChanged(index(0, 0), index(m_maxRow - 1, m_maxCol - 1));
}

// --- 工具和私有方法实现 ---

QString ReportDataModel::cellAddress(int row, int col) const
{
    QString colName;
    int temp = col;
    while (temp >= 0) {
        colName.prepend(QChar('A' + (temp % 26)));
        temp = temp / 26 - 1;
    }
    return colName + QString::number(row + 1);
}

void ReportDataModel::calculateFormula(int row, int col)
{
    RTCell* cell = getCell(row, col);
    if (!cell || !cell->hasFormula)
        return;

    // 调用公式引擎计算结果
    QVariant result = m_formulaEngine->evaluate(cell->formula, this, row, col);
    cell->value = result;
}

RTCell* ReportDataModel::getCell(int row, int col)
{
    return m_cells.value(QPoint(row, col), nullptr);
}

const RTCell* ReportDataModel::getCell(int row, int col) const
{
    return m_cells.value(QPoint(row, col), nullptr);
}

RTCell* ReportDataModel::ensureCell(int row, int col)
{
    QPoint key(row, col);
    if (!m_cells.contains(key)) {
        // 如果单元格不存在，则创建一个新的
        m_cells[key] = new RTCell();
    }
    return m_cells[key];
}


