#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QMainWindow>
#include <QToolBar>
#include <QLineEdit>
#include <QLabel>
#include <QTableView>
#include <QVBoxLayout>
#include <QHBoxLayout>
#include <QAction>
#include <QMenu>

class ReportDataModel;

class MainWindow : public QMainWindow
{
    Q_OBJECT

public:
    MainWindow(QWidget* parent = nullptr);
    ~MainWindow();

private slots:
    void onImportExcel();
    void onExportExcel();
    void onUndo() { /* 待实现 */ }
    void onRedo() { /* 待实现 */ }
    void onFind() { /* 待实现 */ }
    void onFilter() { /* 待实现 */ }
    void onSort() { /* 待实现 */ }
    void onBorderFormat() { /* 待实现 */ }
    void onFontSetting() { /* 待实现 */ }
    void onBackgroundSetting() { /* 待实现 */ }

    void onCurrentCellChanged(const QModelIndex& current, const QModelIndex& previous);
    void onFormulaEditFinished();
    void onCellChanged(int row, int col);

    void onInsertRow();
    void onInsertColumn();
    void onDeleteRow();
    void onDeleteColumn();

private:
    void setupUI();
    void setupToolBar();
    void setupFormulaBar();
    void setupTableView();
    void setupContextMenu();
    void updateFormulaBar(const QModelIndex& index);

private:
    // UI组件
    QWidget* m_centralWidget;
    QVBoxLayout* m_mainLayout;

    // 工具栏
    QToolBar* m_toolBar;

    // 公式栏
    QWidget* m_formulaWidget;
    QLabel* m_cellNameLabel;
    QLineEdit* m_formulaEdit;

    // 表格
    QTableView* m_tableView;
    ReportDataModel* m_dataModel;

    // 右键菜单
    QMenu* m_contextMenu;

    // 状态
    QModelIndex m_currentIndex;
    bool m_updating;
};

#endif // MAINWINDOW_H
