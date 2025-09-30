#ifndef REPORTDATAMODEL_H
#define REPORTDATAMODEL_H

#include "Cell.h"
#include <QHash>
#include <QAbstractTableModel>
#include <QFontInfo>      // 添加这个
#include <QFontDatabase>  // 如果需要的话
#include <QBrush>         // 添加这个
#include <QPoint>
#include <QSize>
#include <QVector> 

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

    QFont ensureFontAvailable(const QFont& requestedFont) const;

    // --- 新增：行高和列宽的接口 ---
    void setRowHeight(int row, double height);
    double getRowHeight(int row) const;
    void setColumnWidth(int col, double width);
    double getColumnWidth(int col) const;
    const QVector<double>& getAllRowHeights() const;
    const QVector<double>& getAllColumnWidths() const;
    void clearSizes();

    // 数据绑定的解析与查询
    void resolveDataBindings();

signals:
    void cellChanged(int row, int col);

    //void dataBingResolving(const QList<QString>& keys);

public:
    RTCell* getCell(int row, int col);
    RTCell* ensureCell(int row, int col);

private:
    QHash<QPoint, RTCell*> m_cells;
    int m_maxRow;
    int m_maxCol;
    FormulaEngine* m_formulaEngine;

    // --- 新增：存储行高和列宽的成员变量 ---
    QVector<double> m_rowHeights;
    QVector<double> m_columnWidths;
};

#endif // REPORTDATAMODEL_H


