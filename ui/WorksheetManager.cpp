#include "WorksheetManager.h"
#include "SpreadsheetView.h"

#include <QInputDialog>
#include <QMessageBox>

WorksheetManager::WorksheetManager(Workbook *workbook, QWidget *parent)
    : QWidget(parent)
    , m_workbook(workbook)
{
    setupUi();

    if (m_workbook) {
        // 为现有工作表创建标签页
        for (int i = 0; i < m_workbook->worksheetCount(); ++i) {
            addWorksheetTab(m_workbook->worksheet(i)->name());
        }
    }
}

void WorksheetManager::setupUi()
{
    auto mainLayout = new QVBoxLayout(this); // 主布局为垂直布局

    // 控制面板
    auto controlGroup = new QGroupBox("工作表控制");
    auto controlLayout = new QHBoxLayout(controlGroup); // 控制面板内部采用水平布局

    // 标签页操作按钮
    m_addTabButton = new QPushButton("添加工作表");
    m_removeTabButton = new QPushButton("删除工作表");

    controlLayout->addWidget(m_addTabButton);
    controlLayout->addWidget(m_removeTabButton);

    controlLayout->addWidget(new QLabel("|")); // 用“|”作分隔符

    // 尺寸设置
    controlLayout->addWidget(new QLabel("行:"));
    m_rowsSpin = new QSpinBox;
    m_rowsSpin->setRange(1, 10000); // 限定范围
    m_rowsSpin->setValue(100); // 默认100行
    controlLayout->addWidget(m_rowsSpin);

    controlLayout->addWidget(new QLabel("列:"));
    m_colsSpin = new QSpinBox;
    m_colsSpin->setRange(1, 1000);
    m_colsSpin->setValue(26);
    controlLayout->addWidget(m_colsSpin);

    m_resizeButton = new QPushButton("应用");
    controlLayout->addWidget(m_resizeButton);

    controlLayout->addWidget(new QLabel("|")); // 分隔符

    // 跳转控制
    controlLayout->addWidget(new QLabel("跳转至Sheet:"));
    m_jumpIndexSpin = new QSpinBox;
    m_jumpIndexSpin->setRange(1, 1); // 初始范围
    controlLayout->addWidget(m_jumpIndexSpin);

    m_jumpButton = new QPushButton("跳转");
    controlLayout->addWidget(m_jumpButton);

    controlLayout->addStretch(); // 右侧填充空间（靠左）

    mainLayout->addWidget(controlGroup);

    // 标签页组件
    m_tabWidget = new QTabWidget; // 标签页控件
    m_tabWidget->setTabsClosable(true); // 启用标签页关闭按钮
    mainLayout->addWidget(m_tabWidget);

    // 连接信号与槽
    connect(m_addTabButton, &QPushButton::clicked, this, &WorksheetManager::onAddTabClicked); // 添加标签页
    connect(m_removeTabButton, &QPushButton::clicked, this, &WorksheetManager::onRemoveTabClicked);// 删除标签页
    connect(m_resizeButton, &QPushButton::clicked, this, &WorksheetManager::onResizeClicked); // 尺寸设置
    connect(m_jumpButton, &QPushButton::clicked, this, &WorksheetManager::onJumpToIndex); // 跳转至

    connect(m_tabWidget, &QTabWidget::currentChanged, this, &WorksheetManager::onTabChanged); // 标签页切换
    connect(m_tabWidget, &QTabWidget::tabCloseRequested, this, [this](int index) {
        if (m_tabWidget->count() > 1) {
            removeCurrentTab();
        }
        else {
            QMessageBox::information(this, "提示", "唯一的工作表，不可删除");
        }
    }); // 关闭标签页（至少保留一个标签页）
}

void WorksheetManager::setWorkbook(Workbook *workbook)
{
    m_workbook = workbook; // 切换工作簿

    // 清空现有标签页与视图
    m_tabWidget->clear();
    m_spreadsheetViews.clear();

    // 重新创建标签页
    if (m_workbook) {
        for (int i = 0; i < m_workbook->worksheetCount(); i++) {
            addWorksheetTab(m_workbook->worksheet(i)->name());
        }
    }
}

SpreadsheetView* WorksheetManager::currentSpreadsheetView() const
{
    int index = m_tabWidget->currentIndex(); // 获取当前标签页索引
    if (index >= 0 && index < m_spreadsheetViews.size()) {
        return m_spreadsheetViews[index]; // 返回对应视图
    }
    return nullptr;
}

// 添加标签页
void WorksheetManager::addWorksheetTab(const QString &name)
{
    QString sheetName = name;
    if (sheetName.isEmpty()) {
        bool ok;
        sheetName = QInputDialog::getText(this, "新工作表",
                                          "请输入工作表名:",
                                          QLineEdit::Normal,
                                          QString("Sheet%1").arg(m_tabWidget->count() + 1),
                                          &ok); // 会自动生成默认格式的名称
        if (!ok || sheetName.isEmpty()) {
            return; // 用户取消或输入空名称时退出
        }
    }

    // 如果工作簿中没有对应的工作表，创建
    if (m_workbook && m_tabWidget->count() >= m_workbook->worksheetCount()) {
        m_workbook->addWorksheet(sheetName);
    }

    // 创建新的视图
    auto spreadsheetView = new SpreadsheetView(m_workbook);
    m_spreadsheetViews.append(spreadsheetView);

    int index = m_tabWidget->addTab(spreadsheetView, sheetName); // 添加到标签页控件
    m_tabWidget->setCurrentIndex(index); // 自动切换到新标签页

    if (m_workbook && index < m_workbook->worksheetCount()) {
        m_workbook->setCurrentWorksheet(index);
    }

    // 更新跳转范围
    m_jumpIndexSpin->setRange(1, m_tabWidget->count());
}

// 删除标签页
void WorksheetManager::removeCurrentTab()
{
    int currentIndex = m_tabWidget->currentIndex();
    if (currentIndex >= 0 && m_tabWidget->count() > 1) { // 确认不是最后一个标签页
        QMessageBox::StandardButton reply = QMessageBox::question(this,
                                                                  "删除工作表",
                                                                  "确定要删除此工作表吗?",
                                                                  QMessageBox::Yes | QMessageBox::No);

        if (reply == QMessageBox::Yes) { // 确认删除
            m_tabWidget->removeTab(currentIndex);

            if (currentIndex < m_spreadsheetViews.size()) {
                // 删除对应视图并从列表中移除
                delete m_spreadsheetViews[currentIndex];
                m_spreadsheetViews.removeAt(currentIndex);
            }

            if (m_workbook) {
                // 从工作簿中删除对应工作表
                m_workbook->removeWorksheet(currentIndex);
            }

            // 更新跳转范围
            m_jumpIndexSpin->setRange(1, m_tabWidget->count());
        }
    }
}

// 标签页切换
void WorksheetManager::onTabChanged(int index)
{
    if (m_workbook) {
        m_workbook->setCurrentWorksheet(index);

        // 刷新当前视图以显示正确的工作表数据
        auto currentView = currentSpreadsheetView();
        if (currentView) {
            currentView->refresh();  // 重新加载当前工作表的数据
        }
    }
    emit currentWorksheetChanged(index);
}

void WorksheetManager::onAddTabClicked()
{
    addWorksheetTab(); // 交给具体方法处理
}

void WorksheetManager::onRemoveTabClicked()
{
    removeCurrentTab();
}

void WorksheetManager::onResizeClicked()
{
    auto currentView = currentSpreadsheetView();
    if (currentView) {
        int rows = m_rowsSpin->value(); // 从QSpinBox中获取设置的行列数
        int cols = m_colsSpin->value();
        currentView->resizeSheet(rows, cols);
    }
}

// 跳转至
void WorksheetManager::onJumpToIndex()
{
    int index = m_jumpIndexSpin->value() - 1; // 转换为由0开始的索引
    if (index >= 0 && index < m_tabWidget->count()) {
        m_tabWidget->setCurrentIndex(index);
    }
}

// 更新标签页名称
void WorksheetManager::updateTabNames()
{
    if (!m_workbook) return;

    for (int i = 0; i < m_tabWidget->count() && i < m_workbook->worksheetCount(); ++i) {
        auto sheet = m_workbook->worksheet(i);
        if (sheet) {
            m_tabWidget->setTabText(i, sheet->name()); // 工作表名称同步给标签页
        }
    }
}
