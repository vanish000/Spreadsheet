#pragma once
#include <QString>
#include <QJsonObject>
#include <QJsonDocument>
#include <QJsonArray>
#include "Workbook.h"

class FileManager
{
public:
    // 文件保存与打开
    static bool saveWorkbook(const Workbook *workbook, const QString &fileName);
    static bool loadWorkbook(Workbook *workbook, const QString &fileName);

    // CSV格式导入和导出
    static bool exportToCsv(const Worksheet *worksheet, const QString &fileName);
    static bool importFromCsv(Worksheet *worksheet, const QString &fileName);

private:
    // JSON转换
    static QJsonObject worksheetToJson(const Worksheet *worksheet);
    static void jsonToWorksheet(const QJsonObject &json, Worksheet *worksheet);

    static QJsonObject cellToJson(const Cell *cell);
    static void jsonToCell(const QJsonObject &json, Cell *cell);
};

