#pragma execution_character_set("utf-8")

#include "mainwindow.h"
#include "reportdatamodel.h"
#include "EnhancedTableView.h"
#include <QApplication>
#include <QFileDialog>
#include <QMessageBox>
#include <QHeaderView>
#include <QLineEdit>
#include <QInputDialog>
#include <QSortFilterProxyModel>

MainWindow::MainWindow(QWidget* parent)
    : QMainWindow(parent)
    , m_updating(false)
	, m_formulaEditMode(false)
    , m_filterModel(nullptr)
{
    setupUI();
    setupToolBar();
    setupFormulaBar();
    setupTableView();
    setupContextMenu();
    setupFindDialog();

    setWindowTitle("SCADA报表控件 v1.0");
    resize(1200, 800);
}

MainWindow::~MainWindow()
{
}

void MainWindow::setupUI()
{
    m_centralWidget = new QWidget(this);
    setCentralWidget(m_centralWidget);

    m_mainLayout = new QVBoxLayout(m_centralWidget);
    m_mainLayout->setSpacing(0);
    m_mainLayout->setContentsMargins(0, 0, 0, 0);
}

void MainWindow::setupToolBar()
{
    m_toolBar = addToolBar("主工具栏");

    // 文件操作
    m_toolBar->addAction("导入", this, &MainWindow::onImportExcel);
    m_toolBar->addAction("导出", this, &MainWindow::onExportExcel);
    m_toolBar->addSeparator();
    m_toolBar->addSeparator();

    // 工具操作
    m_toolBar->addAction("查找", this, &MainWindow::onFind);
    m_toolBar->addAction("筛选", this, &MainWindow::onFilter);
    m_toolBar->addSeparator();

    // 清楚筛选
    m_toolBar->addAction("清除筛选", this, &MainWindow::onClearFilter);
}

void MainWindow::setupFormulaBar()
{
    m_formulaWidget = new QWidget();
    QHBoxLayout* layout = new QHBoxLayout(m_formulaWidget);
    layout->setContentsMargins(5, 5, 5, 5);

    m_cellNameLabel = new QLabel("A1");
    m_cellNameLabel->setMinimumWidth(60);
    m_cellNameLabel->setMaximumWidth(60);
    m_cellNameLabel->setAlignment(Qt::AlignCenter);
    m_cellNameLabel->setStyleSheet("QLabel { border: 1px solid gray; padding: 2px; }");

    m_formulaEdit = new QLineEdit();
    m_formulaEdit->setStyleSheet("QLineEdit { border: 1px solid gray; padding: 2px; }");

    layout->addWidget(m_cellNameLabel);
    layout->addWidget(m_formulaEdit);

    m_mainLayout->addWidget(m_formulaWidget);

    connect(m_formulaEdit, &QLineEdit::editingFinished, this, &MainWindow::onFormulaEditFinished);
    connect(m_formulaEdit, &QLineEdit::textChanged, this, &MainWindow::onFormulaTextChanged);
}

void MainWindow::setupTableView()
{
    m_tableView = new EnhancedTableView();
    m_dataModel = new ReportDataModel(this);

    m_tableView->setModel(m_dataModel);
    m_tableView->setSelectionBehavior(QAbstractItemView::SelectItems);
    m_tableView->setSelectionMode(QAbstractItemView::SingleSelection);
    m_tableView->setContextMenuPolicy(Qt::CustomContextMenu);

    m_tableView->horizontalHeader()->setDefaultSectionSize(80);
    m_tableView->verticalHeader()->setDefaultSectionSize(25);
    m_tableView->setAlternatingRowColors(false);
    m_tableView->setGridStyle(Qt::SolidLine);

    m_mainLayout->addWidget(m_tableView);

    connect(m_tableView->selectionModel(), &QItemSelectionModel::currentChanged,
        this, &MainWindow::onCurrentCellChanged);
    connect(m_dataModel, &ReportDataModel::cellChanged,
        this, &MainWindow::onCellChanged);
    connect(m_tableView, &QTableView::customContextMenuRequested,
        [this](const QPoint& pos) {
            m_contextMenu->exec(m_tableView->mapToGlobal(pos));
        });

    connect(m_tableView, &QTableView::clicked, this, &MainWindow::onCellClicked);
}

void MainWindow::setupContextMenu()
{
    m_contextMenu = new QMenu(this);

    m_contextMenu->addAction("插入行", this, &MainWindow::onInsertRow);
    m_contextMenu->addAction("插入列", this, &MainWindow::onInsertColumn);
    m_contextMenu->addSeparator();
    m_contextMenu->addAction("删除行", this, &MainWindow::onDeleteRow);
    m_contextMenu->addAction("删除列", this, &MainWindow::onDeleteColumn);
}

void MainWindow::setupFindDialog()
{
    m_findDialog = new QDialog(this);
    m_findDialog->setWindowTitle("查找");
    m_findDialog->setModal(false);
    m_findDialog->resize(300, 120);

    QVBoxLayout* layout = new QVBoxLayout(m_findDialog);

    QHBoxLayout* inputLayout = new QHBoxLayout();
    inputLayout->addWidget(new QLabel("查找内容:"));
    m_findLineEdit = new QLineEdit();
    inputLayout->addWidget(m_findLineEdit);
    layout->addLayout(inputLayout);

    QHBoxLayout* buttonLayout = new QHBoxLayout();
    QPushButton* findNextBtn = new QPushButton("查找下一个");
    QPushButton* closeBtn = new QPushButton("关闭");
    buttonLayout->addWidget(findNextBtn);
    buttonLayout->addWidget(closeBtn);
    layout->addLayout(buttonLayout);

    connect(findNextBtn, &QPushButton::clicked, this, &MainWindow::onFindNext);
    connect(closeBtn, &QPushButton::clicked, m_findDialog, &QDialog::hide);
    connect(m_findLineEdit, &QLineEdit::returnPressed, this, &MainWindow::onFindNext);
}

void MainWindow::updateTableSpans()
{
    const auto& allCells = m_dataModel->getAllCells();

    // 清除现有的span设置
    // QTableView没有直接的clearAllSpans方法，需要重置

    // 设置新的span
    for (auto it = allCells.constBegin(); it != allCells.constEnd(); ++it) {
        const QPoint& pos = it.key();
        const RTCell* cell = it.value();

        if (cell && cell->mergedRange.isValid() && cell->mergedRange.isMerged()) {
            // 只为主单元格设置span
            if (pos.x() == cell->mergedRange.startRow &&
                pos.y() == cell->mergedRange.startCol) {

                int rowSpan = cell->mergedRange.rowSpan();
                int colSpan = cell->mergedRange.colSpan();

                if (rowSpan > 1 || colSpan > 1) {
                    m_tableView->setSpan(pos.x(), pos.y(), rowSpan, colSpan);
                }
            }
        }
    }
}

void MainWindow::onImportExcel()
{
    QString fileName = QFileDialog::getOpenFileName(this,
        "导入Excel文件", "", "Excel文件 (*.xlsx *.xls)");

    if (!fileName.isEmpty()) {
        // --- 核心修改：在加载模型数据前，先清除视图的旧状态 ---
        m_tableView->clearSpans(); // 清除旧的合并单元格信息
        // ------------------------------------------------

        if (m_dataModel->loadFromExcel(fileName)) {
            updateTableSpans();
            applyRowColumnSizes();
            QMessageBox::information(this, "成功", "文件导入成功！");
        }
        else {
            QMessageBox::warning(this, "错误", "文件导入失败！");
        }
    }
}

void MainWindow::onExportExcel()
{
    //QString fileName = QFileDialog::getSaveFileName(this,
    //    "导出Excel文件", "", "Excel文件 (*.xlsx)");

    //if (!fileName.isEmpty()) {
    //    if (m_dataModel->saveToExcel(fileName)) {
    //        QMessageBox::information(this, "成功", "文件导出成功！");
    //    }
    //    else {
    //        QMessageBox::warning(this, "错误", "文件导出失败！");
    //    }
    //}
}

void MainWindow::onFind()
{
    if (m_findDialog->isVisible()) {
        m_findDialog->raise();
        m_findDialog->activateWindow();
    }
    else {
        m_findDialog->show();
    }
    m_findLineEdit->setFocus();
    m_findLineEdit->selectAll();
}

void MainWindow::onFindNext()
{
    QString searchText = m_findLineEdit->text();
    if (searchText.isEmpty()) {
        return;
    }

    // 从当前位置开始查找
    QModelIndex startIndex = m_currentIndex.isValid() ? m_currentIndex : m_dataModel->index(0, 0);
    int startRow = startIndex.row();
    int startCol = startIndex.column();

    // 从下一个单元格开始搜索
    bool found = false;
    for (int row = startRow; row < m_dataModel->rowCount() && !found; ++row) {
        int beginCol = (row == startRow) ? startCol + 1 : 0;
        for (int col = beginCol; col < m_dataModel->columnCount() && !found; ++col) {
            QModelIndex index = m_dataModel->index(row, col);
            QString cellText = m_dataModel->data(index, Qt::DisplayRole).toString();

            if (cellText.contains(searchText, Qt::CaseInsensitive)) {
                m_tableView->setCurrentIndex(index);
                m_tableView->scrollTo(index);
                found = true;
            }
        }
    }

    if (!found) {
        // 从头开始搜索到当前位置
        for (int row = 0; row <= startRow && !found; ++row) {
            int endCol = (row == startRow) ? startCol : m_dataModel->columnCount() - 1;
            for (int col = 0; col <= endCol && !found; ++col) {
                QModelIndex index = m_dataModel->index(row, col);
                QString cellText = m_dataModel->data(index, Qt::DisplayRole).toString();

                if (cellText.contains(searchText, Qt::CaseInsensitive)) {
                    m_tableView->setCurrentIndex(index);
                    m_tableView->scrollTo(index);
                    found = true;
                }
            }
        }
    }

    if (!found) {
        QMessageBox::information(this, "查找", "未找到匹配的内容");
    }
}

void MainWindow::onFilter()
{
    bool ok;
    QString filterText = QInputDialog::getText(this, "筛选",
        "请输入筛选条件（支持通配符*和?）:", QLineEdit::Normal, "", &ok);

    if (ok && !filterText.isEmpty()) {
        if (!m_filterModel) {
            m_filterModel = new QSortFilterProxyModel(this);
            m_filterModel->setSourceModel(m_dataModel);
            m_filterModel->setFilterKeyColumn(-1); // 搜索所有列
            m_tableView->setModel(m_filterModel);
        }

        // 转换通配符为正则表达式
        QString regexPattern = QRegularExpression::escape(filterText);
        regexPattern.replace("\\*", ".*");
        regexPattern.replace("\\?", ".");

        m_filterModel->setFilterRegularExpression(QRegularExpression(regexPattern, QRegularExpression::CaseInsensitiveOption));

        //statusBar()->showMessage(QString("筛选结果：%1 行").arg(m_filterModel->rowCount()), 3000);
    }
}

void MainWindow::onClearFilter()
{
    if (m_filterModel) {
        m_tableView->setModel(m_dataModel);
        delete m_filterModel;
        m_filterModel = nullptr;
        //statusBar()->showMessage("已清除筛选", 2000);
    }
}


void MainWindow::onCurrentCellChanged(const QModelIndex& current, const QModelIndex& previous)
{
    Q_UNUSED(previous)

    if (m_updating)
        return;
	// 如果在公式编辑模式下，不更新当前单元格
    if (m_formulaEditMode)
        return;

    m_currentIndex = current;
    updateFormulaBar(current);
}

void MainWindow::onFormulaEditFinished()
{
    if (m_updating || !m_currentIndex.isValid())
        return;

    // 退出公式编辑模式
    if (m_formulaEditMode)
    {
		exitFormulaEditMode();
    }

    QString text = m_formulaEdit->text();
    m_updating = true;
    m_dataModel->setData(m_currentIndex, text, Qt::EditRole);
    m_updating = false;
}

void MainWindow::enterFormulaEditMode()
{
    m_formulaEditMode = true;
    m_formulaEditingIndex = m_currentIndex;

    // 改变公式编辑框的样式，提示用户进入了公式编辑模式
    m_formulaEdit->setStyleSheet("QLineEdit { border: 2px solid blue; padding: 2px; background-color: #f0f8ff; }");

    // 可以在状态栏显示提示信息
    //statusBar()->showMessage("公式编辑模式：点击单元格插入引用，按Enter完成编辑", 5000);
}

void MainWindow::exitFormulaEditMode()
{
    m_formulaEditMode = false;
    m_formulaEditingIndex = QModelIndex();

    // 恢复公式编辑框的正常样式
    m_formulaEdit->setStyleSheet("QLineEdit { border: 1px solid gray; padding: 2px; }");

    //statusBar()->clearMessage();
}

bool MainWindow::isInFormulaEditMode() const
{
    return m_formulaEditMode;
}


void MainWindow::onCellChanged(int row, int col)
{
    if (m_currentIndex.isValid() &&
        m_currentIndex.row() == row &&
        m_currentIndex.column() == col) {
        updateFormulaBar(m_currentIndex);
    }
}

void MainWindow::onCellClicked(const QModelIndex& index)
{
    if (!index.isValid())
    {
        return;
    }
    // 如果在公式编辑状态下，将点击的单元格地址添加到公式中
    if (m_formulaEditMode)
    {
        QString cellAddress = m_dataModel->cellAddress(index.row(), index.column());

        // 获取当前光标位置
        int cursorPos = m_formulaEdit->cursorPosition();
        QString currentText = m_formulaEdit->text();

        // 在光标位置插入单元格地址
        QString newText = currentText.left(cursorPos) + cellAddress + currentText.mid(cursorPos);
        
        m_updating = true;
        m_formulaEdit->setText(newText);
        m_formulaEdit->setCursorPosition(cursorPos + cellAddress.length());
        m_formulaEdit->setFocus(); // 保持焦点在公式编辑框
		m_updating = false;
        
        return;
    }

    // 正常模式下的单元格切换
    if (index != m_currentIndex)
    {
        m_currentIndex = index;
        updateFormulaBar(index);
    }

}

void MainWindow::onFormulaTextChanged()
{
    if (m_updating)
    {
        return;
    }

    QString text = m_formulaEdit->text();

    if (text.startsWith('=') && !m_formulaEditMode)
    {
        enterFormulaEditMode();
    }
    else if (!text.startsWith('=') && m_formulaEditMode)
    {
        exitFormulaEditMode();
    }

}

void MainWindow::updateFormulaBar(const QModelIndex& index)
{
    if (!index.isValid()) {
        m_cellNameLabel->setText("");
        m_formulaEdit->clear();
        return;
    }

    m_updating = true;

	// 更新单元格名称和公式内容
    QString cellName = m_dataModel->cellAddress(index.row(), index.column());
    m_cellNameLabel->setText(cellName);

    //  获取单元格的编辑数据（如果公式则显示公式，否则显示值）
    QVariant editData = m_dataModel->data(index, Qt::EditRole);
    m_formulaEdit->setText(editData.toString());

    // 检查是否应该进入公式编辑模式
    if (editData.toString().startsWith('=')) {
        // 如果新选中的单元格包含公式，但我们不在编辑模式，不要自动进入
        // 只有用户开始编辑时才进入公式编辑模式

        // 可以给公式编辑框一个特殊的样式提示，表明这是一个公式单元格
        m_formulaEdit->setStyleSheet("QLineEdit { border: 1px solid #4CAF50; padding: 2px; background-color: #f9fff9; }");
    }
    else {
        // 普通单元格的样式
        m_formulaEdit->setStyleSheet("QLineEdit { border: 1px solid gray; padding: 2px; }");
    }

    m_updating = false;
}

void MainWindow::onInsertRow()
{
    if (m_currentIndex.isValid()) {
        int row = m_currentIndex.row();
        m_dataModel->insertRows(row, 1);
    }
}

void MainWindow::onInsertColumn()
{
    if (m_currentIndex.isValid()) {
        int col = m_currentIndex.column();
        m_dataModel->insertColumns(col, 1);
    }
}

void MainWindow::onDeleteRow()
{
    if (m_currentIndex.isValid()) {
        int row = m_currentIndex.row();
        m_dataModel->removeRows(row, 1);
    }
}

void MainWindow::onDeleteColumn()
{
    if (m_currentIndex.isValid()) {
        int col = m_currentIndex.column();
        m_dataModel->removeColumns(col, 1);
    }
}

void MainWindow::applyRowColumnSizes()
{
    const auto& colWidths = m_dataModel->getAllColumnWidths();
    for (int i = 0; i < colWidths.size(); ++i) {
        if (colWidths[i] > 0) {
            m_tableView->setColumnWidth(i, static_cast<int>(colWidths[i]));
        }
    }

    const auto& rowHeights = m_dataModel->getAllRowHeights();
    for (int i = 0; i < rowHeights.size(); ++i) {
        if (rowHeights[i] > 0) {
            m_tableView->setRowHeight(i, static_cast<int>(rowHeights[i]));
        }
    }
}
