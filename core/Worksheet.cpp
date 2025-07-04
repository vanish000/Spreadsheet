#include "Worksheet.h"

Worksheet::Worksheet(const QString &name, QObject *parent)
    : QObject(parent)
    , m_name(name)
    , m_rowCount(100)
    , m_colCount(26)
{}

void Worksheet::setName(const QString &name)
{
    if (m_name != name) {
        m_name = name;
        emit nameChanged(name);
    }
}

// 单元格访问
std::shared_ptr<Cell> Worksheet::cell(int row, int col)
{
    QPair<int, int> key(row, col);

    if (!m_cells.contains(key)) { // 单元格不存在则创建
        auto newCell = std::make_shared<Cell>(row, col, this);
        m_cells[key] = newCell;

        // 信号槽连接：cell对象发送单元格内容改变信号，从而触发工作表的单元格改变信号
        connect(newCell.get(), &Cell::valueChanged,
                this, [this, row, col]() { emit cellChanged(row, col); });
        // 捕获列表[this, row, col]中，this捕获当前对象，允许调用成员函数。
    }

    return m_cells[key];
}


void Worksheet::setCell(int row, int col, std::shared_ptr<Cell> cell)
{
    QPair<int, int> key(row, col);
    m_cells[key] = cell; // 传入单元格覆盖指定位置。
    emit cellChanged(row, col);
}

// 插入or删除行列暂未实现
void Worksheet::insertRow(int row)
{
    m_rowCount++;
}

void Worksheet::insertColumn(int col)
{
    m_colCount++;
}

void Worksheet::removeRow(int row)
{
    if (m_rowCount > 1) {
        m_rowCount--;
    }
}

void Worksheet::removeColumn(int col)
{
    if (m_colCount > 1) {
        m_colCount--;
    }
}

void Worksheet::clear()
{
    m_cells.clear(); // 移除所有单元格
}
