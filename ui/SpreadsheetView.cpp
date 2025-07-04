#include "SpreadsheetView.h"
#include "CellDetailEditor.h"
#include "../core/Cell.h"

#include <QDebug>


SpreadsheetView::SpreadsheetView(Workbook *workbook, QWidget *parent)
    : QTableWidget(parent)
    , m_workbook(workbook)
{
    // 初始化
    setRowCount(100);
    setColumnCount(26);

    setupHeaders();

    if (m_workbook) {
        loadData();
    }

    // 连接信号槽
    connect(this, &QTableWidget::cellChanged,
            this, &SpreadsheetView::onCellChanged);
    connect(this, &QTableWidget::currentCellChanged,
            this, &SpreadsheetView::onCurrentCellChanged);
}

void SpreadsheetView::mouseDoubleClickEvent(QMouseEvent *event)
{
    if (event->button() == Qt::LeftButton) { // 只处理左键双击事件
        editCellDetails(); // 打开编辑器
    }
    else {
        QTableWidget::mouseDoubleClickEvent(event); // 其余事件仍由基类处理
    }
}

void SpreadsheetView::editCellDetails()
{
    int row = currentRow();
    int col = currentColumn();

    if (row >= 0 && col >= 0 && m_workbook && m_workbook->currentWorksheet()) {
        auto cell = m_workbook->currentWorksheet()->cell(row, col);

        CellDetailEditor editor(cell, this); // 创建编辑器，父窗口为当前视图
        if (editor.exec() == QDialog::Accepted) { // 应用更改
            blockSignals(true); // 阻塞信号，防止初次创建时
            auto item = this->item(row, col);
            if (!item) {
                item = new QTableWidgetItem();
                setItem(row, col, item);
            }
            item->setText(cell->displayText()); // 显示同步
            blockSignals(false);
        }
    }
}

void SpreadsheetView::resizeSheet(int rows, int cols)
{
    setRowCount(rows);
    setColumnCount(cols);
    setupHeaders();
    loadData();
}

void SpreadsheetView::setWorkbook(Workbook *workbook)
{
    m_workbook = workbook;
    loadData();
}

// 提供手动更新数据的公有接口
void SpreadsheetView::refresh()
{
    loadData();
}

void SpreadsheetView::setupHeaders()
{
    // 设置列标题 (A, B, C, ...)
    QStringList horizontalLabels;
    for (int i = 0; i < columnCount(); ++i) {
        QString label;
        int n = i;
        while(n >= 0){
            label.prepend(QChar('A' + (n % 26)));
            n = n / 26 - 1;
        }
        horizontalLabels << label;
    }
    setHorizontalHeaderLabels(horizontalLabels);

    // 设置行标题 (1, 2, 3, ...)
    QStringList verticalLabels;
    for (int i = 0; i < rowCount(); i++) {
        verticalLabels << QString::number(i + 1);
    }
    setVerticalHeaderLabels(verticalLabels);

    // 设置默认列宽、行高
    horizontalHeader()->setDefaultSectionSize(80);
    verticalHeader()->setDefaultSectionSize(25);
}

void SpreadsheetView::loadData()
{
    // 安全验证：确认工作簿和工作表存在
    if (!m_workbook || !m_workbook->currentWorksheet()) {
        return;
    }

    auto sheet = m_workbook->currentWorksheet();

    // 阻止信号发送，以避免触发cellChanged导致递归
    blockSignals(true);

    clearContents();

    for (int row = 0; row < rowCount(); row++) {
        for (int col = 0; col < columnCount(); col++) {
            auto cell = sheet->cell(row, col);
            if (cell && (!cell->displayText().isEmpty() || !cell->value().isNull())) {
                // 为有内容的单元格创建item并显示文本
                setItem(row, col, new QTableWidgetItem(cell->displayText()));
            }
        }
    }

    // 数据载入完成，恢复信号
    blockSignals(false);
}

void SpreadsheetView::onCellChanged(int row, int column)
{
    // 安全验证
    if (!m_workbook || !m_workbook->currentWorksheet()) {
        return;
    }

    auto sheet = m_workbook->currentWorksheet();
    auto cell = sheet->cell(row, column);

    auto item = this->item(row, column);
    if (item) {
        QString text = item->text();
        if (text.startsWith("=")) { // 开头为“=”，识别为公式
            // 处理公式
            cell->setFormula(text);
            item->setText(cell->displayText());
        }
        else {
            cell->setValue(text); // 直接显示内容
        }
    }
}

void SpreadsheetView::onCurrentCellChanged(int currentRow, int currentColumn,
                                           int previousRow, int previousColumn)
{
    // 暂不处理前一单元格，留待扩展功能（如保存编辑历史）
    Q_UNUSED(previousRow)
    Q_UNUSED(previousColumn)

    if (m_workbook && m_workbook->currentWorksheet()) {
        auto cell = m_workbook->currentWorksheet()->cell(currentRow, currentColumn);
        // 可以显示单元格的公式或其他信息
    }
}
