#pragma once

#include <QObject>
#include <QList>
#include <memory>

#include "Worksheet.h"

class Workbook : public QObject
{
    Q_OBJECT

public:
    explicit Workbook(QObject *parent = nullptr);

    // 计数
    int worksheetCount() const { return m_worksheets.size(); }

    // 工作表访问
    std::shared_ptr<Worksheet> worksheet(int index) const;
    std::shared_ptr<Worksheet> currentWorksheet() const;

    // 工作表管理
    void addWorksheet(const QString &name = QString());
    void removeWorksheet(int index);
    void setCurrentWorksheet(int index);

    // 文件IO
    bool saveToFile(const QString &fileName);
    bool loadFromFile(const QString &fileName);

signals:
    void worksheetAdded(int index);
    void worksheetRemoved(int index);
    void currentWorksheetChanged(int index);

private:
    QList<std::shared_ptr<Worksheet>> m_worksheets; // 工作表列表
    int m_currentIndex; // 当前活动工作表索引
};
