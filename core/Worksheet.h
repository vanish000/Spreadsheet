#pragma once

#include <QObject>
#include <QHash>
#include <memory>

#include "Cell.h"

class Worksheet : public QObject
{
    Q_OBJECT

public:
    explicit Worksheet(const QString &name = "Sheet1", QObject *parent = nullptr);

    // 名称访问与设置
    QString name() const { return m_name; }
    void setName(const QString &name);

    // 单元格访问与设置
    std::shared_ptr<Cell> cell(int row, int col);
    void setCell(int row, int col, std::shared_ptr<Cell> cell);

    // 表格尺寸查询
    int rowCount() const { return m_rowCount; }
    int columnCount() const { return m_colCount; }

    // 表格操作
    void insertRow(int row);
    void insertColumn(int col);
    void removeRow(int row);
    void removeColumn(int col);

    void clear();

signals:
    void cellChanged(int row, int col);
    void nameChanged(const QString &name);

private:
    QString m_name; // 工作表名称
    QHash<QPair<int, int>, std::shared_ptr<Cell>> m_cells; // 哈希表存储单元格，只为有数据的单元格分配内存，减少开销
    int m_rowCount;
    int m_colCount; // 行列数
};
