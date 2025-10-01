#include "EnhancedTableView.h"
#include "reportdatamodel.h"
#include "DataBindingConfig.h"
#include <QPaintEvent>
#include <QPainter>
#include <QHeaderView>
#include <QAbstractProxyModel>
#include <QPen>

EnhancedTableView::EnhancedTableView(QWidget* parent)
    : QTableView(parent)
{
}

void EnhancedTableView::paintEvent(QPaintEvent* event)
{
    // �ȵ��û���Ļ���
    QTableView::paintEvent(event);

    // Ȼ������Զ���߿�
    QPainter painter(viewport());
    drawBorders(&painter);
}

void EnhancedTableView::drawBorders(QPainter* painter)
{
    ReportDataModel* reportModel = getReportModel();
    if (!reportModel) return;

    painter->save();

    // ��ȡ�ɼ�����
    QRect viewportRect = viewport()->rect();
    int firstRow = rowAt(viewportRect.top());
    int lastRow = rowAt(viewportRect.bottom());
    int firstCol = columnAt(viewportRect.left());
    int lastCol = columnAt(viewportRect.right());

    if (firstRow < 0) firstRow = 0;
    if (lastRow < 0) lastRow = reportModel->rowCount() - 1;
    if (firstCol < 0) firstCol = 0;
    if (lastCol < 0) lastCol = reportModel->columnCount() - 1;

    // ����ÿ����Ԫ��ı߿�
    for (int row = firstRow; row <= lastRow; ++row) {
        for (int col = firstCol; col <= lastCol; ++col) {
            const CellData* cell = reportModel->getCell(row, col);
            if (!cell) continue;

            QModelIndex index = reportModel->index(row, col);
            QRect cellRect = visualRect(index);

            // ����ϲ���Ԫ��
            QSize spanSize = reportModel->span(index);
            if (spanSize.width() > 1 || spanSize.height() > 1) {
                // ����ϲ���Ԫ�����������
                for (int r = 1; r < spanSize.height(); ++r) {
                    QModelIndex spanIndex = reportModel->index(row + r, col);
                    cellRect = cellRect.united(visualRect(spanIndex));
                }
                for (int c = 1; c < spanSize.width(); ++c) {
                    QModelIndex spanIndex = reportModel->index(row, col + c);
                    cellRect = cellRect.united(visualRect(spanIndex));
                }
                for (int r = 1; r < spanSize.height(); ++r) {
                    for (int c = 1; c < spanSize.width(); ++c) {
                        QModelIndex spanIndex = reportModel->index(row + r, col + c);
                        cellRect = cellRect.united(visualRect(spanIndex));
                    }
                }
            }

            // ���Ʊ߿�
            const RTCellBorder& border = cell->style.border;

            // ��߿�
            if (border.left != RTBorderStyle::None) {
                QPen pen(border.leftColor);
                pen.setWidth(static_cast<int>(border.left));
                painter->setPen(pen);
                painter->drawLine(cellRect.topLeft(), cellRect.bottomLeft());
            }

            // �ұ߿�
            if (border.right != RTBorderStyle::None) {
                QPen pen(border.rightColor);
                pen.setWidth(static_cast<int>(border.right));
                painter->setPen(pen);
                painter->drawLine(cellRect.topRight(), cellRect.bottomRight());
            }

            // �ϱ߿�
            if (border.top != RTBorderStyle::None) {
                QPen pen(border.topColor);
                pen.setWidth(static_cast<int>(border.top));
                painter->setPen(pen);
                painter->drawLine(cellRect.topLeft(), cellRect.topRight());
            }

            // �±߿�
            if (border.bottom != RTBorderStyle::None) {
                QPen pen(border.bottomColor);
                pen.setWidth(static_cast<int>(border.bottom));
                painter->setPen(pen);
                painter->drawLine(cellRect.bottomLeft(), cellRect.bottomRight());
            }
        }
    }

    painter->restore();
}

void EnhancedTableView::setSpan(int row, int column, int rowSpanCount, int columnSpanCount)
{
    QTableView::setSpan(row, column, rowSpanCount, columnSpanCount);
}

void EnhancedTableView::updateSpans()
{
    ReportDataModel* reportModel = getReportModel();
    if (!reportModel) return;

    // ������еĺϲ���Ԫ������
    clearSpans();

    // �����������кϲ���Ԫ��
    const auto& allCells = reportModel->getAllCells();
    for (auto it = allCells.constBegin(); it != allCells.constEnd(); ++it) {
        const QPoint& pos = it.key();
        const CellData* cell = it.value();

        if (cell && cell->isMergedMain()) {
            int rowSpan = cell->mergedRange.rowSpan();
            int colSpan = cell->mergedRange.colSpan();
            if (rowSpan > 1 || colSpan > 1) {
                setSpan(pos.x(), pos.y(), rowSpan, colSpan);
            }
        }
    }
}

ReportDataModel* EnhancedTableView::getReportModel() const
{
    // ������ܴ��ڵĴ���ģ��
    QAbstractItemModel* currentModel = model();

    // ����Ǵ���ģ�ͣ���ȡԴģ��
    while (QAbstractProxyModel* proxyModel = qobject_cast<QAbstractProxyModel*>(currentModel)) {
        currentModel = proxyModel->sourceModel();
    }

    return qobject_cast<ReportDataModel*>(currentModel);
}