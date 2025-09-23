#pragma execution_character_set("utf-8")


#include "mainwindow.h"
#include "reportdatamodel.h"
#include <QApplication>
#include <QFileDialog>
#include <QMessageBox>
#include <QHeaderView>

MainWindow::MainWindow(QWidget* parent)
    : QMainWindow(parent)
    , m_updating(false)
	, m_formulaEditMode(false)
{
    setupUI();
    setupToolBar();
    setupFormulaBar();
    setupTableView();
    setupContextMenu();

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

    // 编辑操作
    m_toolBar->addAction("撤销", this, &MainWindow::onUndo);
    m_toolBar->addAction("重做", this, &MainWindow::onRedo);
    m_toolBar->addSeparator();

    // 工具操作
    m_toolBar->addAction("查找", this, &MainWindow::onFind);
    m_toolBar->addAction("筛选", this, &MainWindow::onFilter);
    m_toolBar->addAction("排序", this, &MainWindow::onSort);
    m_toolBar->addSeparator();

    // 格式操作
    m_toolBar->addAction("边框", this, &MainWindow::onBorderFormat);
    m_toolBar->addAction("字体", this, &MainWindow::onFontSetting);
    m_toolBar->addAction("背景", this, &MainWindow::onBackgroundSetting);
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
    m_tableView = new QTableView();
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

void MainWindow::onImportExcel()
{
    QString fileName = QFileDialog::getOpenFileName(this,
        "导入Excel文件", "", "Excel文件 (*.xlsx *.xls)");

    if (!fileName.isEmpty()) {
        if (m_dataModel->loadFromExcel(fileName)) {
            QMessageBox::information(this, "成功", "文件导入成功！");
        }
        else {
            QMessageBox::warning(this, "错误", "文件导入失败！");
        }
    }
}

void MainWindow::onExportExcel()
{
    QString fileName = QFileDialog::getSaveFileName(this,
        "导出Excel文件", "", "Excel文件 (*.xlsx)");

    if (!fileName.isEmpty()) {
        if (m_dataModel->saveToExcel(fileName)) {
            QMessageBox::information(this, "成功", "文件导出成功！");
        }
        else {
            QMessageBox::warning(this, "错误", "文件导出失败！");
        }
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
