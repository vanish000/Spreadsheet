#include "Cell.h"
#include <QRegularExpression>

Cell::Cell(int row, int col, QObject *parent)
    : QObject(parent)
    , m_row(row)
    , m_col(col)
    , m_readOnly(false)
{}

void Cell::setValue(const QVariant &value)
{
    if (m_value != value) { // 检查值变化
        m_value = value;
        m_formula.clear();
        emit valueChanged(); // 发射信号
    }
}

void Cell::setFormula(const QString &formula)
{
    if (m_formula != formula) {
        m_formula = formula;
        evaluateFormula(); // 立即根据公式求值
        emit formulaChanged();
    }
}

QString Cell::displayText() const
{
    // 优先显示公式
    if (!m_formula.isEmpty()) {
        return m_formula;
    }
    return m_value.toString();
}

void Cell::evaluateFormula()
{
    if (m_formula.isEmpty()) {
        return;
    }

    // 简单的公式计算 (仅支持基本数学运算)
    QString expr = m_formula;
    if (expr.startsWith("=")) {
        expr = expr.mid(1); // 从下标1截取，去掉“ = ”

        // 简单的数学表达式计算
        QRegularExpression re(R"((\d+(?:\.\d+)?)\s*([+\-*/])\s*(\d+(?:\.\d+)?))");
        QRegularExpressionMatch match = re.match(expr); // 执行正则匹配
        // 正则表达式：捕获组匹配数字 "(\d+(?:\.\d+)?)" （"\d+"1个或多个数字，"(?:\.\d+)?"可选的小数部分）
        // \s*：零个或多个空白字符
        // 捕获组匹配运算符 "([+\-*/])"
        // 使用 R"(...)" 表示采用原始字符串，从而无需使用转义字符

        if (match.hasMatch()) {
            // 获取捕获结果
            double num1 = match.captured(1).toDouble();
            QString op = match.captured(2);
            double num2 = match.captured(3).toDouble();

            // 公式计算
            double result = 0;
            switch (op.at(0).toLatin1()) {
                case '+': result = num1 + num2; break;
                case '-': result = num1 - num2; break;
                case '*': result = num1 * num2; break;
                case '/': result = (num2 != 0) ? num1 / num2 : 0; break;
                default:  result = 0; break;
            }
            m_value = result;
        }
        else {
            m_value = 0; // 无效公式
        }
    }

    emit valueChanged();
}
