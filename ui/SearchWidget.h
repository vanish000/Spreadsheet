#pragma once

#include <QWidget>
#include <QLineEdit>
#include <QPushButton>
#include <QLabel>
#include <QCheckBox>
#include <QHBoxLayout>
#include <QGroupBox>


class SpreadsheetView;

class SearchWidget : public QWidget
{
    Q_OBJECT

public:
    explicit SearchWidget(QWidget *parent = nullptr);

    void setSpreadsheetView(SpreadsheetView *view); // 重设新视图

signals:
    void cellFound(int row, int col);

private slots:
    void onSearchClicked();
    void onNextClicked();
    void onPreviousClicked();
    void onReplaceClicked();

private:
    void setupUi();
    void performSearch(); // 执行查找
    void jumpToResult(int index); // 跳转至结果

    SpreadsheetView *m_spreadsheetView;

    QLineEdit *m_searchEdit; // 接收查找关键词
    QLineEdit *m_replaceEdit; // 接收替换文本
    QPushButton *m_searchButton;
    QPushButton *m_nextButton;
    QPushButton *m_previousButton;
    QPushButton *m_replaceButton;
    QPushButton *m_replaceAllButton;

    QCheckBox *m_caseSensitiveCheck; // 区分大小写
    QCheckBox *m_wholeWordCheck; // 全词匹配

    QLabel *m_resultLabel;

    struct SearchResult { //匹配结果
        int row;
        int col;
        QString value;
    };

    QList<SearchResult> m_searchResults;
    int m_currentResultIndex;
};

