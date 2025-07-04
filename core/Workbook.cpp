#include "Workbook.h"

#include <QFile> // 文件操作类
#include <QTextStream> // 文本流，便于格式化文本读写


Workbook::Workbook(QObject *parent)
    : QObject(parent)
    , m_currentIndex(-1)
{
    addWorksheet("Sheet1"); // 初始工作表
}

// 访问工作表
std::shared_ptr<Worksheet> Workbook::worksheet(int index) const
{
    if (index >= 0 && index < m_worksheets.size()) {
        return m_worksheets[index];
    }
    return nullptr;
}

// 访问当前工作表
std::shared_ptr<Worksheet> Workbook::currentWorksheet() const
{
    return worksheet(m_currentIndex);
}

// 添加工作表
void Workbook::addWorksheet(const QString &name)
{
    // 生成工作表默认名称
    QString sheetName = name.isEmpty() ?
                        QString("Sheet%1").arg(m_worksheets.size() + 1) : name;

    // 创建工作表对象，并加入列表
    auto sheet = std::make_shared<Worksheet>(sheetName, this);
    m_worksheets.append(sheet);

    // 如果尚未设置当前工作表，则将第一个表设置为当前
    if (m_currentIndex == -1) {
        m_currentIndex = 0;
    }

    // 发送信号：新增工作表，索引为 size() - 1
    emit worksheetAdded(m_worksheets.size() - 1);
}

// 删除工作表
void Workbook::removeWorksheet(int index)
{
    // 至少保留一个工作表防止工作簿为空
    if (index >= 0 && index < m_worksheets.size() && m_worksheets.size() > 1) {
        m_worksheets.removeAt(index);

        if (m_currentIndex >= m_worksheets.size()) { // 当前活动工作表索引越界
            m_currentIndex = m_worksheets.size() - 1;
        }

        // 发送信号：删除工作表，索引为 index
        emit worksheetRemoved(index);
    }
}

// 设置活动工作表
void Workbook::setCurrentWorksheet(int index)
{
    // 传入相同索引，则不改变
    if (index >= 0 && index < m_worksheets.size() && index != m_currentIndex) {
        m_currentIndex = index;
        // 发出信号：当前工作表变更
        emit currentWorksheetChanged(index);
    }
}

// 保存文件（未完全实现）
bool Workbook::saveToFile(const QString &fileName)
{
    QFile file(fileName);
    if (!file.open(QIODevice::WriteOnly)) { // 以“只写”方式打开文件，如果文件不存在则创建文件
        return false; // 打开失败（被占用、没有写入权限等）
    }

    // 文件内容写入（暂时只实现了写入基本信息，待完整的序列化）
    QTextStream stream(&file); // 创建文本流
    stream << "Spreadsheet File\n";
    stream << "Worksheets: " << m_worksheets.size() << "\n";

    return true;
}

// 文件加载（未完全实现）
bool Workbook::loadFromFile(const QString &fileName)
{
    QFile file(fileName);
    if (!file.open(QIODevice::ReadOnly)) { // 以“只读”模式打开文件
        return false;
    }

    // 清空现有工作表
    m_worksheets.clear();
    addWorksheet("Sheet1");

    return true;
}
