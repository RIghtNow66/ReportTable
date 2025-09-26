#ifndef REPORTDATAMODEL_H
#define REPORTDATAMODEL_H

#include "Cell.h"
#include <QHash>
#include <QAbstractTableModel>
#include <QPoint>
#include <QSize>

inline uint qHash(const QPoint& key, uint seed = 0) noexcept
{
    // 用 QPair<int,int> 的哈希来实现
    return qHash(QPair<int, int>(key.x(), key.y()), seed);
}

struct RTCell;
struct RTCellStyle;
class FormulaEngine;

class ReportDataModel : public QAbstractTableModel
{
    Q_OBJECT

public:
    explicit ReportDataModel(QObject* parent = nullptr);
    ~ReportDataModel();

    // --- Qt Model 接口 ---
    int rowCount(const QModelIndex& parent = QModelIndex()) const override;
    int columnCount(const QModelIndex& parent = QModelIndex()) const override;
    QVariant data(const QModelIndex& index, int role = Qt::DisplayRole) const override;
    bool setData(const QModelIndex& index, const QVariant& value, int role = Qt::EditRole) override;
    Qt::ItemFlags flags(const QModelIndex& index) const override;
    QVariant headerData(int section, Qt::Orientation orientation, int role = Qt::DisplayRole) const override;
    
    // 支持合并单元格
    QSize span(const QModelIndex& index) const;
    
    bool insertRows(int row, int count, const QModelIndex& parent = QModelIndex()) override;
    bool removeRows(int row, int count, const QModelIndex& parent = QModelIndex()) override;
    bool insertColumns(int column, int count, const QModelIndex& parent = QModelIndex()) override;
    bool removeColumns(int column, int count, const QModelIndex& parent = QModelIndex()) override;

    // --- 文件操作 (现在会调用ExcelHandler) ---
    bool loadFromExcel(const QString& fileName);
    bool saveToExcel(const QString& fileName);

    // --- 公共接口 (供ExcelHandler和FormulaEngine使用) ---
    void clearAllCells();
    void addCellDirect(int row, int col, RTCell* cell);
    void updateModelSize(int newRowCount, int newColCount);
    const QHash<QPoint, RTCell*>& getAllCells() const;
    void recalculateAllFormulas();
    const RTCell* getCell(int row, int col) const;
    void calculateFormula(int row, int col);
    QString cellAddress(int row, int col) const;

signals:
    void cellChanged(int row, int col);

public:
    RTCell* getCell(int row, int col);
    RTCell* ensureCell(int row, int col);

private:
    QHash<QPoint, RTCell*> m_cells;
    int m_maxRow;
    int m_maxCol;
    FormulaEngine* m_formulaEngine;
};

#endif // REPORTDATAMODEL_H
