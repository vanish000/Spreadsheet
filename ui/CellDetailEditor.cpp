#include "CellDetailEditor.h"
#include "../core/Cell.h"  // 包含Cell类的完整定义

#include <QDebug>

CellDetailEditor::CellDetailEditor(std::shared_ptr<Cell> cell, QWidget *parent)
    : QDialog(parent)
    , m_cell(cell)
    , m_cellAddressEdit(nullptr)
    , m_valueEdit(nullptr)
    , m_formulaEdit(nullptr)
    , m_previewLabel(nullptr)
    , m_okButton(nullptr)
    , m_cancelButton(nullptr)
    , m_clearButton(nullptr)
{
    // 初始化
    setupUi();
    loadCellData();
    connectSignals();
}

void CellDetailEditor::setupUi()
{
    setWindowTitle("单元格编辑");
    setModal(true); // 阻塞父窗口，关闭对话框才能进行其他操作
    resize(500, 400);

    // 设置窗口图标和属性
    setWindowFlags(windowFlags() & ~Qt::WindowContextHelpButtonHint); // 移除标题栏的“?”帮助按钮
    // Qt的窗口标志是枚举值，实际上是位掩码，通过 & ~ 的位操作即可移除特定按钮

    auto mainLayout = new QVBoxLayout(this); // 垂直布局
    mainLayout->setSpacing(10); // 间距10像素
    mainLayout->setContentsMargins(10, 10, 10, 10); // 对话框边缘留10像素

    // 单元格地址
    auto addressGroup = new QGroupBox("单元格地址"); // 分组框，并设置标题
    auto addressLayout = new QFormLayout(addressGroup); // 表单布局

    m_cellAddressEdit = new QLineEdit;
    m_cellAddressEdit->setReadOnly(true); // 设为只读
    m_cellAddressEdit->setStyleSheet("QLineEdit { background-color: #f5f5f5; }"); // CSS样式设置：灰色背景
    addressLayout->addRow("地址:", m_cellAddressEdit); // 添加 标签-控件对 到box中

    mainLayout->addWidget(addressGroup);

    // 值编辑区域
    auto valueGroup = new QGroupBox("数值");
    auto valueLayout = new QFormLayout(valueGroup);

    m_valueEdit = new QLineEdit;
    m_valueEdit->setPlaceholderText("输入单元格的值"); // 在输入框中添加提示语句
    valueLayout->addRow("值:", m_valueEdit);

    mainLayout->addWidget(valueGroup);

    // 公式编辑区域
    auto formulaGroup = new QGroupBox("公式");
    auto formulaLayout = new QVBoxLayout(formulaGroup);

    m_formulaEdit = new QTextEdit; // 多行文本编辑器
    m_formulaEdit->setMaximumHeight(100); // 高度限制
    m_formulaEdit->setPlaceholderText("输入公式 (如: =A1+B1)");
    formulaLayout->addWidget(m_formulaEdit);

    mainLayout->addWidget(formulaGroup);

    // 显示预览区域
    auto previewGroup = new QGroupBox("预览");
    auto previewLayout = new QVBoxLayout(previewGroup);

    m_previewLabel = new QLabel("此处为预览显示");
    m_previewLabel->setStyleSheet(
        "QLabel { "
        "background-color: #f0f0f0; "
        "padding: 10px; "
        "border: 1px solid #ccc; "
        "border-radius: 4px; "
        "min-height: 60px; "
        "}"
        ); // CSS样式：定义背景、边框、圆角等
    m_previewLabel->setWordWrap(true); // 长文本自动换行
    m_previewLabel->setAlignment(Qt::AlignTop | Qt::AlignLeft); // 左上角对齐
    previewLayout->addWidget(m_previewLabel);

    mainLayout->addWidget(previewGroup);

    // 控制台
    auto buttonLayout = new QHBoxLayout;
    buttonLayout->setSpacing(10);

    m_clearButton = new QPushButton("清除");
    m_clearButton->setToolTip("清除所有输入内容"); // 悬停时显示说明

    m_cancelButton = new QPushButton("取消");
    m_cancelButton->setToolTip("取消编辑并关闭");

    m_okButton = new QPushButton("确定");
    m_okButton->setToolTip("保存更改并关闭");
    m_okButton->setDefault(true); // 默认设置：按回车键自动触发

    // 设置按钮样式
    m_okButton->setStyleSheet("QPushButton { font-weight: bold; }"); // “确认”按钮样式：粗体

    buttonLayout->addWidget(m_clearButton);
    buttonLayout->addStretch(); // 在清除按钮与其他按钮之间填充空间
    buttonLayout->addWidget(m_cancelButton);
    buttonLayout->addWidget(m_okButton);

    mainLayout->addLayout(buttonLayout);
}

void CellDetailEditor::connectSignals()
{
    // 文本变化
    connect(m_valueEdit, &QLineEdit::textChanged,
            this, &CellDetailEditor::onValueChanged);
    connect(m_formulaEdit, &QTextEdit::textChanged,
            this, &CellDetailEditor::onFormulaChanged);

    // 按钮触发
    connect(m_okButton, &QPushButton::clicked,
            this, &CellDetailEditor::accept); // 确认：重载的accept，执行一系列额外的输入验证后再关闭对话框
    connect(m_cancelButton, &QPushButton::clicked,
            this, &QDialog::reject); // 取消：直接调用reject，关闭对话框
    connect(m_clearButton, &QPushButton::clicked,
            this, &CellDetailEditor::onClearClicked); // 清除：调用自定义的槽函数
}

void CellDetailEditor::loadCellData()
{
    if (!m_cell) {
        qWarning() << "CellDetailEditor: Cell pointer is null";
        return;
    }

    try {
        // 显示单元格地址
        QString address = QString("%1%2")
                          .arg(QChar('A' + m_cell->column()))
                          .arg(m_cell->row() + 1);
        m_cellAddressEdit->setText(address);

        // 加载值和公式（优先显示公式）
        QString formula = m_cell->formula();
        if (!formula.isEmpty()) {
            m_formulaEdit->setPlainText(formula);
            m_valueEdit->clear(); // 互斥显示
        }
        else {
            QString value = m_cell->value().toString();
            m_valueEdit->setText(value);
            m_formulaEdit->clear();
        }

        updatePreview(); // 更新预览

    } catch (const std::exception& e) { // 捕获异常
        qWarning() << "Error loading cell data:" << e.what();
    }
}

void CellDetailEditor::updatePreview()
{
    // 获取更新
    QString preview;
    QString formulaText = m_formulaEdit->toPlainText().trimmed(); // 多行文本需使用toPlainText提取
    QString valueText = m_valueEdit->text().trimmed(); // trimmed去除首尾可能存在的冗余空白字符

    if (!formulaText.isEmpty()) {
        preview = QString("公式: %1\n结果: %2")
                  .arg(formulaText)
                  .arg("(计算结果将在应用后显示)");
    }
    else if (!valueText.isEmpty()) {
        preview = QString("值: %1").arg(valueText);
    }
    else {
        preview = "空单元格";
    }

    m_previewLabel->setText(preview);
}

void CellDetailEditor::onFormulaChanged()
{
    QString formulaText = m_formulaEdit->toPlainText().trimmed();

    // 输入公式后，清空值输入框
    if (!formulaText.isEmpty()) {
        // 临时断开信号连接，避免递归调用
        m_valueEdit->blockSignals(true);
        m_valueEdit->clear();
        m_valueEdit->blockSignals(false);
    }

    updatePreview();
}

void CellDetailEditor::onValueChanged()
{
    QString valueText = m_valueEdit->text().trimmed();

    // 输入值后，清空公式输入框
    if (!valueText.isEmpty()) {
        // 临时断开信号连接，避免递归调用
        m_formulaEdit->blockSignals(true);
        m_formulaEdit->clear();
        m_formulaEdit->blockSignals(false);
    }

    updatePreview();
}

void CellDetailEditor::onClearClicked()
{
    // 临时断开信号，避免不必要的更新
    m_valueEdit->blockSignals(true);
    m_formulaEdit->blockSignals(true);

    m_valueEdit->clear();
    m_formulaEdit->clear();

    // 重新连接信号
    m_valueEdit->blockSignals(false);
    m_formulaEdit->blockSignals(false);

    updatePreview();

    // 设置焦点到值输入框
    m_valueEdit->setFocus();
}

QString CellDetailEditor::getValue() const
{
    return m_valueEdit->text().trimmed();
}

QString CellDetailEditor::getFormula() const
{
    return m_formulaEdit->toPlainText().trimmed();
}

void CellDetailEditor::accept()
{
    if (!m_cell) {
        qWarning() << "CellDetailEditor: Cannot save - cell pointer is null";
        QDialog::reject();
        return;
    }

    try {
        QString formula = getFormula();
        QString value = getValue();

        // 验证输入
        if (!formula.isEmpty()) {
            // 验证公式格式
            if (!formula.startsWith('=')) {
                formula = '=' + formula;  // 自动添加等号
            }
            m_cell->setFormula(formula);
        }
        else if (!value.isEmpty()) {
            m_cell->setValue(value);
        }
        else {
            // 清空单元格
            m_cell->setValue(QVariant());
        }

        qDebug() << "Cell updated successfully";
        QDialog::accept(); // 调用基类方法

    } catch (const std::exception& e) {
        qWarning() << "Error saving cell data:" << e.what();
        QDialog::reject();
    }
}
