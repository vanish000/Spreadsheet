#pragma once
#include <QTableWidget> // 表格控件
#include <QHeaderView> // 表头视图，用于自定义行列表头
#include <QMouseEvent> // 鼠标事件类，处理鼠标双击等操作

#include "../core/Workbook.h"

class SpreadsheetView : public QTableWidget
{
    Q_OBJECT

public:
    explicit SpreadsheetView(Workbook *workbook = nullptr, QWidget *parent = nullptr);

    void setWorkbook(Workbook *workbook); // 重置当前工作簿
    void refresh(); // 刷新视图显示
    void resizeSheet(int rows, int cols); // 调整表格尺寸

protected:
    void mouseDoubleClickEvent(QMouseEvent *event) override; // 自定义鼠标双击行为

private slots:
    void onCellChanged(int row, int column);
    void onCurrentCellChanged(int currentRow, int currentColumn,
                              int previousRow, int previousColumn);
    void editCellDetails(); // 打开单元格内容编辑界面

private:
    void setupHeaders(); // 设置表头
    void loadData(); // 视图载入数据

    Workbook *m_workbook; // 指向当前工作簿
};
