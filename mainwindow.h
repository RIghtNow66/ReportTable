﻿#ifndef MAINWINDOW_H
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
#include <QDialog>
#include <QPushButton>
#include <QSortFilterProxyModel>

class ReportDataModel;

class MainWindow : public QMainWindow
{
    Q_OBJECT

public:
    MainWindow(QWidget* parent = nullptr);
    ~MainWindow();

private slots:
    // 文件操作
    void onImportExcel();
    void onExportExcel();

    // 工具操作
    void onFind();
    void onFindNext();
    void onFilter();
    void onClearFilter();

    void onCurrentCellChanged(const QModelIndex& current, const QModelIndex& previous);
    void onFormulaEditFinished();
    void onFormulaTextChanged(); // 新增：监听公式编辑变化
    void onCellClicked(const QModelIndex& index); // 新增：处理单元格点击
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
    void setupFindDialog();
    void updateFormulaBar(const QModelIndex& index);

    void enterFormulaEditMode();
    void exitFormulaEditMode();
    bool isInFormulaEditMode() const;
    void updateTableSpans();

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
    QSortFilterProxyModel* m_filterModel;

    // 查找对话框
    QDialog* m_findDialog;
    QLineEdit* m_findLineEdit;

    // 右键菜单
    QMenu* m_contextMenu;

    // 状态
    QModelIndex m_currentIndex;
    bool m_updating;
    bool m_formulaEditMode; // 新增：标识是否在公式编辑模式
    QModelIndex m_formulaEditingIndex; // 新增：记录正在编辑公式的单元格
};

#endif // MAINWINDOW_H
