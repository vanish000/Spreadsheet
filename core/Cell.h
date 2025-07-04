#pragma once

#include <QObject>
#include <QVariant> // 通用数据容器
#include <QString>

class Cell : public QObject
{
    Q_OBJECT

public:
    explicit Cell(int row = 0, int col = 0, QObject *parent = nullptr);

    // 单元格内容处理
    QVariant value() const { return m_value; } // 获取内容数据：支持数字、字符串、日期等多种数据类型
    void setValue(const QVariant &value);

    QString formula() const { return m_formula; } // 获取原始公式
    void setFormula(const QString &formula);

    QString displayText() const; // 显示文本接口

    // 行列坐标接口
    int row() const { return m_row; }
    int column() const { return m_col; }

    // 访问控制
    bool isReadOnly() const { return m_readOnly; }
    void setReadOnly(bool readOnly) { m_readOnly = readOnly; }

signals:
    void valueChanged();
    void formulaChanged();

private:
    QVariant m_value;
    QString m_formula; // 原始公式
    int m_row;
    int m_col;
    bool m_readOnly;

    void evaluateFormula(); // 公式处理
};
