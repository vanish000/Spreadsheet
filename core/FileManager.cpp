#include "FileManager.h"
#include "Cell.h"
#include "Worksheet.h"

#include <QFile> // 文件读写
#include <QTextStream> // 格式化文件读写
#include <QDebug>

// 保存
bool FileManager::saveWorkbook(const Workbook *workbook, const QString &fileName)
{
    if (!workbook) return false;

    // JSON根对象（解析和处理 JSON 的入口）
    QJsonObject rootObject;
    rootObject["version"] = "1.0";
    rootObject["application"] = "Spreadsheet";

    // 工作表转换为JSON并构建数组
    QJsonArray worksheetsArray;
    for (int i = 0; i < workbook->worksheetCount(); ++i) {
        auto worksheet = workbook->worksheet(i);
        if (worksheet) {
            worksheetsArray.append(worksheetToJson(worksheet.get())); // 使用get获取原始指针
        }
    }

    rootObject["worksheets"] = worksheetsArray;
    rootObject["currentWorksheet"] = 0; // 默认设为第一个工作表

    QJsonDocument doc(rootObject); // 封装为可序列化文档

    // 写入文件
    QFile file(fileName); // 创建文件对象
    if (!file.open(QIODevice::WriteOnly)) {
        qDebug() << "Failed to open file for writing:" << fileName;
        return false;
    }

    file.write(doc.toJson()); // JSON序列化，写入文件对象
    return true;
}

// 打开文件
bool FileManager::loadWorkbook(Workbook *workbook, const QString &fileName)
{
    if (!workbook) return false;

    QFile file(fileName);
    if (!file.open(QIODevice::ReadOnly)) {
        qDebug() << "Failed to open file for reading:" << fileName;
        return false;
    }

    // JSON解析
    QByteArray data = file.readAll();
    QJsonDocument doc = QJsonDocument::fromJson(data); // 解析为JSON文档

    if (doc.isNull() || !doc.isObject()) { // 检查有效和是否为对象
        qDebug() << "Invalid JSON format";
        return false;
    }

    QJsonObject rootObject = doc.object(); // 获取根对象

    // 检查版本兼容性
    QString version = rootObject["version"].toString();
    if (version != "1.0") {
        qDebug() << "Unsupported file version:" << version;
        return false;
    }

    // 清空现有工作表
    while (workbook->worksheetCount() > 0) {
        workbook->removeWorksheet(0);
    }

    // 加载工作表
    QJsonArray worksheetsArray = rootObject["worksheets"].toArray(); // 获取工作表数组
    for (const QJsonValue &value : worksheetsArray) {
        if (value.isObject()) {
            // 工作表创建
            QJsonObject worksheetObj = value.toObject();
            QString name = worksheetObj["name"].toString();

            workbook->addWorksheet(name);
            auto worksheet = workbook->worksheet(workbook->worksheetCount() - 1);
            if (worksheet) {
                jsonToWorksheet(worksheetObj, worksheet.get()); // 数据转换，逆序列化
            }
        }
    }

    // 没有工作表，则创建一个默认工作表
    if (workbook->worksheetCount() == 0) {
        workbook->addWorksheet("Sheet1");
    }

    return true;
}

// 导出为CSV
bool FileManager::exportToCsv(const Worksheet *worksheet, const QString &fileName)
{
    if (!worksheet) return false;

    QFile file(fileName);
    if (!file.open(QIODevice::WriteOnly | QIODevice::Text)) { //以只写文本模式（自动处理换行符）打开
        return false;
    }

    QTextStream stream(&file); // 创建文本流，绑定到file上

    // 确定实际行列范围
    int maxRow = 0, maxCol = 0;
    for (int row = 0; row < worksheet->rowCount(); ++row) {
        for (int col = 0; col < worksheet->columnCount(); ++col) {
            auto cell = const_cast<Worksheet*>(worksheet)->cell(row, col); // 临时移除const，避免编译错误（cell方法没有const版本）
            if (cell && (!cell->value().isNull() || !cell->formula().isEmpty())) {
                maxRow = qMax(maxRow, row);
                maxCol = qMax(maxCol, col); // 获取实际使用的最大行列号
            }
        }
    }

    // 导出数据
    for (int row = 0; row <= maxRow; ++row) {
        QStringList rowData; // 保存当前行每一格的数据
        for (int col = 0; col <= maxCol; ++col) {
            auto cell = const_cast<Worksheet*>(worksheet)->cell(row, col);
            QString cellValue;
            if (cell) {
                cellValue = cell->value().toString(); // 将单元格值转换为字符串
                // CSV格式：如果单元格内容包含逗号或引号，需要用双引号整个包围
                if (cellValue.contains(',') || cellValue.contains('"') || cellValue.contains('\n')) {
                    cellValue.replace('"', "\"\""); // 双引号转义为两个双引号
                    cellValue = '"' + cellValue + '"';
                }
            }
            rowData << cellValue; // 加入行数据
        }
        stream << rowData.join(',') << '\n'; // 每个单元格数据以逗号作为字段分隔符连接成一行写入流中
    }

    return true;
}

// CSV导入
bool FileManager::importFromCsv(Worksheet *worksheet, const QString &fileName)
{
    if (!worksheet) return false;

    QFile file(fileName);
    if (!file.open(QIODevice::ReadOnly | QIODevice::Text)) {
        return false;
    }

    QTextStream stream(&file); // 创建文本流
    int row = 0;

    while (!stream.atEnd()) { // 逐行读取
        QString line = stream.readLine(); // 保存当前行内容
        QStringList fields; // 保存一行解析出的字段

        // CSV解析（处理引号和逗号）
        bool inQuotes = false; // 标记当前是否在引号内
        QString currentField;

        for (int i = 0; i < line.length(); ++i) { // 遍历当前行每个字符
            QChar c = line[i];

            if (c == '"') { // 遇到引号
                if (inQuotes && i + 1 < line.length() && line[i + 1] == '"') { // 如果在引号内，且下一个字符仍为引号，说明为转义的双引号
                    currentField += '"';
                    ++i; // 跳过下一个引号
                }
                else { // 遇到左引号则为true，右引号则为false
                    inQuotes = !inQuotes;
                }
            }
            else if (c == ',' && !inQuotes) { // 遇到逗号，且不在引号内，说明为字段分隔符
                fields << currentField; // 当前字段为一个单元格内容，加入
                currentField.clear();
            }
            else { // 其他字符，直接加入当前字段
                currentField += c;
            }
        }
        fields << currentField; // 单独将最后一个字段加入（结尾没有逗号）

        // 将数据写入工作表
        for (int col = 0; col < fields.size(); ++col) {
            auto cell = worksheet->cell(row, col);
            cell->setValue(fields[col]);
        }

        ++row;
    }

    return true;
}

// 工作表转换为JSON
QJsonObject FileManager::worksheetToJson(const Worksheet *worksheet)
{
    // 保存基本信息
    QJsonObject worksheetObj;
    worksheetObj["name"] = worksheet->name();
    worksheetObj["rowCount"] = worksheet->rowCount();
    worksheetObj["columnCount"] = worksheet->columnCount();

    QJsonArray cellsArray;

    // 只保存非空单元格
    for (int row = 0; row < worksheet->rowCount(); ++row) {
        for (int col = 0; col < worksheet->columnCount(); ++col) {
            auto cell = const_cast<Worksheet*>(worksheet)->cell(row, col);
            if (cell && (!cell->value().isNull() || !cell->formula().isEmpty())) {
                QJsonObject cellObj = cellToJson(cell.get());
                cellObj["row"] = row;
                cellObj["column"] = col;
                cellsArray.append(cellObj);
            }
        }
    }

    worksheetObj["cells"] = cellsArray;
    return worksheetObj;
}

// JSON转换为工作表
void FileManager::jsonToWorksheet(const QJsonObject &json, Worksheet *worksheet)
{
    worksheet->setName(json["name"].toString());

    // 调整工作表大小
    int rowCount = json["rowCount"].toInt();
    int colCount = json["columnCount"].toInt();

    // resize方法

    // 加载单元格数据
    QJsonArray cellsArray = json["cells"].toArray();
    for (const QJsonValue &value : cellsArray) {
        if (value.isObject()) {
            QJsonObject cellObj = value.toObject();
            int row = cellObj["row"].toInt();
            int col = cellObj["column"].toInt();

            auto cell = worksheet->cell(row, col);
            jsonToCell(cellObj, cell.get());
        }
    }
}

// 单元格转为JSON
QJsonObject FileManager::cellToJson(const Cell *cell)
{
    QJsonObject cellObj;

    if (!cell->formula().isEmpty()) {
        cellObj["formula"] = cell->formula();
    }

    cellObj["value"] = QJsonValue::fromVariant(cell->value());
    cellObj["readOnly"] = cell->isReadOnly();

    return cellObj;
}

// JSON转为单元格
void FileManager::jsonToCell(const QJsonObject &json, Cell *cell)
{
    if (json.contains("formula")) {
        cell->setFormula(json["formula"].toString());
    }
    else {
        cell->setValue(json["value"].toVariant());
    }

    cell->setReadOnly(json["readOnly"].toBool());
}
