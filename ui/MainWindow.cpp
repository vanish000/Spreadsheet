#include "MainWindow.h"
#include "WorksheetManager.h"
#include "SearchWidget.h"
#include "SpreadsheetView.h"
#include "../core/FileManager.h"

#include <QApplication>
#include <QFileDialog>
#include <QMessageBox>
#include <QVBoxLayout>
#include <QMenuBar>
#include <QToolBar>
#include <QStatusBar>
#include <QCloseEvent>

MainWindow::MainWindow(QWidget *parent)
    : QMainWindow(parent)
    , m_worksheetManager(nullptr)
    , m_searchWidget(nullptr)
    , m_workbook(std::make_unique<Workbook>())
    , m_isModified(false)
{
    // 初始化
    setupUi();
    setupMenus();
    setupToolbars();
    connectSignals();

    setCurrentFile("");
    statusBar()->showMessage("就绪", 2000);
}

MainWindow::~MainWindow()
{
}

void MainWindow::setupUi()
{
    setWindowTitle("Spreadsheet"); // 初始标题
    resize(1200, 800); // 默认尺寸

    // 创建中央控件，默认选中和自动适应窗口
    auto centralWidget = new QWidget;
    setCentralWidget(centralWidget);

    auto layout = new QVBoxLayout(centralWidget); // 创建垂直布局

    // 搜索组件
    m_searchWidget = new SearchWidget;
    layout->addWidget(m_searchWidget);

    // 工作表管理器
    m_worksheetManager = new WorksheetManager(m_workbook.get());
    layout->addWidget(m_worksheetManager);

    // 连接搜索组件和当前视图
    auto currentView = m_worksheetManager->currentSpreadsheetView();
    if (currentView) {
        m_searchWidget->setSpreadsheetView(currentView);
    }
}

void MainWindow::setupMenus()
{
    // 文件菜单
    auto fileMenu = menuBar()->addMenu("文件(&F)");
    // 动作创建：动作"新建"加入fileMenu，并与对应槽函数连接
    auto newAction = fileMenu->addAction("新建(&N)", this, &MainWindow::newFile);
    newAction->setShortcut(QKeySequence::New); // 设置快捷键：使用Qt预设的快捷键：Ctrl+N

    auto openAction = fileMenu->addAction("打开(&O)...", this, &MainWindow::openFile);
    openAction->setShortcut(QKeySequence::Open);

    fileMenu->addSeparator(); // 添加分隔线

    auto saveAction = fileMenu->addAction("保存(&S)", this, &MainWindow::saveFile);
    saveAction->setShortcut(QKeySequence::Save);

    auto saveAsAction = fileMenu->addAction("另存为(&A)...", this, &MainWindow::saveAsFile);
    saveAsAction->setShortcut(QKeySequence::SaveAs);

    fileMenu->addSeparator();

    auto exportCsvAction = fileMenu->addAction("导出CSV...", this, &MainWindow::exportToCsv);
    auto importCsvAction = fileMenu->addAction("导入CSV...", this, &MainWindow::importFromCsv);

    fileMenu->addSeparator();

    auto exitAction = fileMenu->addAction("关闭(&X)", this, &QWidget::close);
    exitAction->setShortcut(QKeySequence::Quit);

    // 编辑菜单
    auto editMenu = menuBar()->addMenu("编辑(&E)");
    auto undoAction = editMenu->addAction("撤销(&U)");
    undoAction->setShortcut(QKeySequence::Undo);
    undoAction->setEnabled(false); // 暂未实现，禁用

    auto redoAction = editMenu->addAction("重做(&R)");
    redoAction->setShortcut(QKeySequence::Redo);
    redoAction->setEnabled(false); // 暂时禁用

    editMenu->addSeparator();

    auto cutAction = editMenu->addAction("剪切(&T)");
    cutAction->setShortcut(QKeySequence::Cut);
    cutAction->setEnabled(false); // 暂时禁用

    auto copyAction = editMenu->addAction("复制(&C)");
    copyAction->setShortcut(QKeySequence::Copy);
    copyAction->setEnabled(false); // 暂时禁用

    auto pasteAction = editMenu->addAction("粘贴(&P)");
    pasteAction->setShortcut(QKeySequence::Paste);
    pasteAction->setEnabled(false); // 暂时禁用

    // 视图菜单
    auto viewMenu = menuBar()->addMenu("视图(&V)");
    auto fullscreenAction = viewMenu->addAction("全屏(&F)", this, [this]() {
        if (isFullScreen()) {
            showNormal();
        }
        else {
            showFullScreen();
        }
    });
    fullscreenAction->setShortcut(QKeySequence::FullScreen);

    // 帮助菜单
    auto helpMenu = menuBar()->addMenu("帮助(&H)");
    helpMenu->addAction("关于(&A)", this, &MainWindow::about);
}

void MainWindow::setupToolbars()
{
    // 创建工具栏，添加常用功能的按钮
    auto fileToolbar = addToolBar("文件");
    fileToolbar->addAction("新建", this, &MainWindow::newFile);
    fileToolbar->addAction("打开", this, &MainWindow::openFile);
    fileToolbar->addAction("保存", this, &MainWindow::saveFile);

    fileToolbar->addSeparator();

    fileToolbar->addAction("导出CSV", this, &MainWindow::exportToCsv);
    fileToolbar->addAction("导入CSV", this, &MainWindow::importFromCsv);
}

void MainWindow::connectSignals()
{
    // 连接工作表管理器信号
    connect(m_worksheetManager, &WorksheetManager::currentWorksheetChanged,
            this, &MainWindow::onCurrentWorksheetChanged);
}

void MainWindow::closeEvent(QCloseEvent *event)
{
    if (maybeSave()) { // 关闭前确认是否需要保存
        event->accept();
    }
    else {
        event->ignore();
    }
}

void MainWindow::newFile()
{
    if (maybeSave()) { // 新建前检查当前文件是否需要保存
        m_workbook = std::make_unique<Workbook>(); // 创建新的工作簿实例
        m_worksheetManager->setWorkbook(m_workbook.get()); // 更新工作表管理器

        auto currentView = m_worksheetManager->currentSpreadsheetView();
        if (currentView) {
            m_searchWidget->setSpreadsheetView(currentView); // 更新搜索组件
        }

        setCurrentFile(""); // 清空文件名（路径）
        m_isModified = false; // 重置修改标志
    }
}

void MainWindow::openFile()
{
    if (maybeSave()) {
        // 文件选择对话框，返回文件路径，第4个参数是过滤器：只显示特定扩展名的文件
        QString fileName = QFileDialog::getOpenFileName(this, "打开文件", "",
                                                        "Spreadsheet Files (*.xlsx *.csv *.ssp);;All Files (*)");

        if (!fileName.isEmpty()) {
            if (FileManager::loadWorkbook(m_workbook.get(), fileName)) { // 调用文件管理器的加载方法
                // 交给工作表管理器并设置视图
                m_worksheetManager->setWorkbook(m_workbook.get());

                auto currentView = m_worksheetManager->currentSpreadsheetView();
                if (currentView) {
                    m_searchWidget->setSpreadsheetView(currentView);
                }

                setCurrentFile(fileName);
                m_isModified = false; // 重置修改状态
                statusBar()->showMessage("文件打开成功", 2000);
            }
            else {
                QMessageBox::warning(this, "错误",
                                     QString("打开该文件失败: %1").arg(fileName));
            }
        }
    }
}

void MainWindow::saveFile()
{
    // 首次保存调用"另存为"
    if (m_currentFileName.isEmpty()) {
        saveAsFile();
    }
    else {
        // 保存到原文件中
        if (FileManager::saveWorkbook(m_workbook.get(), m_currentFileName)) {
            m_isModified = false;
            statusBar()->showMessage("文件保存成功", 2000);
        }
        else {
            QMessageBox::warning(this, "错误",
                                 QString("文件保存失败: %1").arg(m_currentFileName));
        }
    }
}

void MainWindow::saveAsFile()
{
    QString fileName = QFileDialog::getSaveFileName(this,"保存文件","",
                                                    "CSV Files (*.csv);;Excel Files (*.xlsx);;SSP Files (*.ssp);;All Files (*)");

    if (!fileName.isEmpty()) {
        if (FileManager::saveWorkbook(m_workbook.get(), fileName)) {
            setCurrentFile(fileName);
            m_isModified = false;
            statusBar()->showMessage("文件保存成功", 2000);
        }
        else {
            QMessageBox::warning(this, "错误",
                                 QString("文件保存失败: %1").arg(fileName));
        }
    }
}

// 导出为CSV
void MainWindow::exportToCsv()
{
    auto worksheet = m_workbook->currentWorksheet();
    if (!worksheet) {
        QMessageBox::information(this, "提示", "没有可导出的工作表");
        return;
    }

    QString fileName = QFileDialog::getSaveFileName(this,
                                                    "导出CSV", worksheet->name() + ".csv",
                                                    "CSV Files (*.csv);;All Files (*)");

    if (!fileName.isEmpty()) {
        if (FileManager::exportToCsv(worksheet.get(), fileName)) { // 调用管理器的导出方法
            statusBar()->showMessage("成功导出为CSV文件", 2000);
        }
        else {
            QMessageBox::warning(this, "错误",
                                 QString("导出CSV文件失败: %1").arg(fileName));
        }
    }
}

// 导入CSV
void MainWindow::importFromCsv()
{
    // 选择要导入的CSV文件
    QString fileName = QFileDialog::getOpenFileName(this,
                                                    "导入CSV", "",
                                                    "CSV Files (*.csv);;All Files (*)");

    if (!fileName.isEmpty()) {
        auto worksheet = m_workbook->currentWorksheet();
        if (worksheet && FileManager::importFromCsv(worksheet.get(), fileName)) {
            auto currentView = m_worksheetManager->currentSpreadsheetView();
            if (currentView) {
                currentView->refresh();
            }
            m_isModified = true; // 导入后即标记为已修改
            statusBar()->showMessage("导入CSV文件成功", 2000);
        } else {
            QMessageBox::warning(this, "错误",
                                 QString("导入CSV失败: %1").arg(fileName));
        }
    }
}

// 工作表切换
void MainWindow::onCurrentWorksheetChanged(int index)
{
    Q_UNUSED(index) // 不使用index（连接的信号会传递index，必须接收）

    // 更新搜索组件的目标视图
    auto currentView = m_worksheetManager->currentSpreadsheetView();
    if (currentView) {
        m_searchWidget->setSpreadsheetView(currentView);
    }

    m_isModified = true;
}

// 确认保存
bool MainWindow::maybeSave()
{
    if (m_isModified) {
        QMessageBox::StandardButton result = QMessageBox::question(this,
                                                                   "Spreadsheet",
                                                                   "此文档已被修改.\n是否要保存这些更改?",
                                                                   QMessageBox::Save | QMessageBox::Discard | QMessageBox::Cancel);

        if (result == QMessageBox::Save) { // 保存
            saveFile();
            return !m_isModified; // 如果保存失败，m_isModified仍为true，此处返回false
        }
        else if (result == QMessageBox::Cancel) { // 取消
            return false;
        }
    }
    return true; // 放弃更改
}

// 设置当前文件名
void MainWindow::setCurrentFile(const QString &fileName)
{
    m_currentFileName = fileName;

    QString title = "Spreadsheet";
    if (!fileName.isEmpty()) {
        QFileInfo fileInfo(fileName);
        title += " - " + fileInfo.fileName(); // 只提取文件名，不显示路径
    }
    else {
        title += " - 新建文档";
    }

    setWindowTitle(title);
}

void MainWindow::about()
{
    QMessageBox::about(this, "About Spreadsheet",
                       "Spreadsheet v1.0.0\n"
                       "A basic spreadsheet application built with Qt6\n"
                       "Developer: Jiang");
}
