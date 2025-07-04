#pragma once

#include <QWidget>
#include <QTabWidget>
#include <QPushButton>
#include <QLineEdit>
#include <QSpinBox>
#include <QLabel>
#include <QHBoxLayout>
#include <QVBoxLayout>
#include <QGroupBox>
#include "../core/Workbook.h"

class SpreadsheetView; // 前向声明（在源文件中再引用）


class WorksheetManager : public QWidget
{
    Q_OBJECT

public:
    explicit WorksheetManager(Workbook *workbook, QWidget *parent = nullptr);

    void setWorkbook(Workbook *workbook); // 切换活动工作簿
    SpreadsheetView* currentSpreadsheetView() const; // 当前视图获取

    void addWorksheetTab(const QString &name = QString()); // 添加标签页
    void removeCurrentTab(); // 移除标签页

signals:
    void currentWorksheetChanged(int index);

private slots:
    void onTabChanged(int index);
    void onAddTabClicked();
    void onRemoveTabClicked();
    void onResizeClicked();
    void onJumpToIndex();

private:
    void setupUi();
    void updateTabNames();

    Workbook *m_workbook; // 当前工作簿
    QTabWidget *m_tabWidget; // 标签页控件：管理多个工作表视图

    // 控制按钮
    QPushButton *m_addTabButton;
    QPushButton *m_removeTabButton;
    QPushButton *m_resizeButton;
    QPushButton *m_jumpButton;

    // 尺寸设置
    QSpinBox *m_rowsSpin; // 整数数值输入控件（可直接输入，也可使用上下箭头调整数值）
    QSpinBox *m_colsSpin;

    // 跳转
    QSpinBox *m_jumpIndexSpin;

    // 视图管理
    QList<SpreadsheetView*> m_spreadsheetViews;
};
