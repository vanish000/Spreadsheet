#include "SearchWidget.h"
#include "SpreadsheetView.h"
#include <QMessageBox>

SearchWidget::SearchWidget(QWidget *parent)
    : QWidget(parent)
    , m_spreadsheetView(nullptr)
    , m_currentResultIndex(-1)
{
    setupUi();
}

void SearchWidget::setupUi()
{
    auto mainLayout = new QHBoxLayout(this); // 主布局设为水平布局

    auto searchGroup = new QGroupBox("查找&替换");
    auto searchLayout = new QHBoxLayout(searchGroup); // 分组框内部也设为水平布局

    // 查找输入
    searchLayout->addWidget(new QLabel("查找:"));
    m_searchEdit = new QLineEdit;
    m_searchEdit->setPlaceholderText("输入查找内容..."); // 占位符
    searchLayout->addWidget(m_searchEdit);

    // 替换输入
    searchLayout->addWidget(new QLabel("替换:"));
    m_replaceEdit = new QLineEdit;
    m_replaceEdit->setPlaceholderText("替换为...");
    searchLayout->addWidget(m_replaceEdit);

    // 按钮
    m_searchButton = new QPushButton("查找");
    m_nextButton = new QPushButton("下一个");
    m_previousButton = new QPushButton("上一个");
    m_replaceButton = new QPushButton("替换");
    m_replaceAllButton = new QPushButton("全部替换");

    searchLayout->addWidget(m_searchButton);
    searchLayout->addWidget(m_nextButton);
    searchLayout->addWidget(m_previousButton);
    searchLayout->addWidget(m_replaceButton);
    searchLayout->addWidget(m_replaceAllButton);

    // 选项
    m_caseSensitiveCheck = new QCheckBox("区分大小写");
    m_wholeWordCheck = new QCheckBox("全词匹配");

    searchLayout->addWidget(m_caseSensitiveCheck);
    searchLayout->addWidget(m_wholeWordCheck);

    // 结果显示
    m_resultLabel = new QLabel("待查找");
    searchLayout->addWidget(m_resultLabel);

    mainLayout->addWidget(searchGroup);

    // 连接信号
    connect(m_searchEdit, &QLineEdit::returnPressed, this, &SearchWidget::onSearchClicked);
    connect(m_searchButton, &QPushButton::clicked, this, &SearchWidget::onSearchClicked);
    connect(m_nextButton, &QPushButton::clicked, this, &SearchWidget::onNextClicked);
    connect(m_previousButton, &QPushButton::clicked, this, &SearchWidget::onPreviousClicked);
    connect(m_replaceButton, &QPushButton::clicked, this, &SearchWidget::onReplaceClicked);

    // 初始状态
    m_nextButton->setEnabled(false);
    m_previousButton->setEnabled(false);
    m_replaceButton->setEnabled(false);
    m_replaceAllButton->setEnabled(false);
}

// 更换视图
void SearchWidget::setSpreadsheetView(SpreadsheetView *view)
{
    m_spreadsheetView = view;
}

void SearchWidget::onSearchClicked()
{
    performSearch();
}

// 查找功能实现
void SearchWidget::performSearch()
{
    if (!m_spreadsheetView || m_searchEdit->text().isEmpty()) {
        return;
    }

    m_searchResults.clear();
    m_currentResultIndex = -1;

    QString searchText = m_searchEdit->text(); // 获取用户输入
    Qt::CaseSensitivity caseSensitivity = m_caseSensitiveCheck->isChecked() ?
                                          Qt::CaseSensitive : Qt::CaseInsensitive; // 是否区分大小写

    // 搜索所有单元格
    for (int row = 0; row < m_spreadsheetView->rowCount(); ++row) {
        for (int col = 0; col < m_spreadsheetView->columnCount(); ++col) {
            auto item = m_spreadsheetView->item(row, col);
            if (item) { // 只检查非空单元格
                QString cellText = item->text();
                bool found = false;

                if (m_wholeWordCheck->isChecked()) {
                    // 全词匹配（查找完整单词，而非某个单词的一部分）
                    // 分割单词：正则表达式\\W+，匹配一个或多个非单词字符（包括空格、标点等），Qt::SkipEmptyParts跳过空字符串
                    QStringList words = cellText.split(QRegularExpression("\\W+"), Qt::SkipEmptyParts);
                    for (const QString &word : words) { // 逐个比较
                        if (word.compare(searchText, caseSensitivity) == 0) {
                            found = true;
                            break;
                        }
                    }
                }
                else {
                    // 部分匹配
                    found = cellText.contains(searchText, caseSensitivity);
                }

                if (found) { // 将查找结果添加到列表中
                    SearchResult result;
                    result.row = row;
                    result.col = col;
                    result.value = cellText;
                    m_searchResults.append(result);
                }
            }
        }
    }

    // 更新UI状态
    if (m_searchResults.isEmpty()) { // 无结果
        m_resultLabel->setText("找不到该内容");
        m_nextButton->setEnabled(false);
        m_previousButton->setEnabled(false);
        m_replaceButton->setEnabled(false);
        m_replaceAllButton->setEnabled(false);
    }
    else {
        m_resultLabel->setText(QString("匹配到 %1 个结果").arg(m_searchResults.size()));
        m_nextButton->setEnabled(true);
        m_previousButton->setEnabled(true);
        m_replaceButton->setEnabled(true);
        m_replaceAllButton->setEnabled(true);

        // 跳转到第一个结果
        m_currentResultIndex = 0;
        jumpToResult(0);
    }
}

// 跳转
void SearchWidget::jumpToResult(int index)
{
    if (index >= 0 && index < m_searchResults.size()) {
        const auto &result = m_searchResults[index];
        m_spreadsheetView->setCurrentCell(result.row, result.col); // 视图移动到目标单元格位置
        m_spreadsheetView->scrollToItem(m_spreadsheetView->item(result.row, result.col));

        m_resultLabel->setText(QString("第 %1 个（共 %2 个）").arg(index + 1).arg(m_searchResults.size()));
        emit cellFound(result.row, result.col);
    }
}

// 下一个结果
void SearchWidget::onNextClicked()
{
    if (!m_searchResults.isEmpty()) {
        m_currentResultIndex = (m_currentResultIndex + 1) % m_searchResults.size();
        jumpToResult(m_currentResultIndex);
    }
}

// 上一个结果
void SearchWidget::onPreviousClicked()
{
    if (!m_searchResults.isEmpty()) {
        m_currentResultIndex = (m_currentResultIndex - 1 + m_searchResults.size()) % m_searchResults.size();
        jumpToResult(m_currentResultIndex);
    }
}

// 替换
void SearchWidget::onReplaceClicked()
{
    if (m_currentResultIndex >= 0 && m_currentResultIndex < m_searchResults.size()) {
        const auto &result = m_searchResults[m_currentResultIndex];
        auto item = m_spreadsheetView->item(result.row, result.col);
        if (item) { // 单元格存在
            QString newText = item->text();
            QString searchText = m_searchEdit->text();
            QString replaceText = m_replaceEdit->text();

            Qt::CaseSensitivity caseSensitivity = m_caseSensitiveCheck->isChecked() ?
                                                  Qt::CaseSensitive : Qt::CaseInsensitive;

            if (m_wholeWordCheck->isChecked()) {
                // 全词替换（确保替换的是完整单词而非一部分）
                // 正则表达式：\\b是单词边界，确保捕获完整单词
                newText.replace(QRegularExpression(QString("\\b%1\\b").arg(QRegularExpression::escape(searchText)),
                                caseSensitivity == Qt::CaseInsensitive ? QRegularExpression::CaseInsensitiveOption : QRegularExpression::NoPatternOption),
                                replaceText); // 此处对caseSensitivity还进行了类型转换，不可直接使用caseSensitivity
            }
            else {
                // 部分替换
                newText.replace(searchText, replaceText, caseSensitivity);
            }

            item->setText(newText);

            // 更新搜索结果
            performSearch();
        }
    }
}
