#pragma once

#include <QMainWindow> // 主窗口基类，提供菜单栏等功能
#include <QTabWidget> // 标签页控件，支持多个工作表
#include <QMenuBar>
#include <QToolBar>
#include <QStatusBar> // 界面组件
#include <QVBoxLayout>
#include <memory>

#include "../core/Workbook.h"

class WorksheetManager;
class SearchWidget;

class MainWindow : public QMainWindow
{
    Q_OBJECT

public:
    MainWindow(QWidget *parent = nullptr);
    ~MainWindow();

protected:
    void closeEvent(QCloseEvent* event) override;

private slots:
    // 文件操作
    void newFile();
    void openFile();
    void saveFile();
    void saveAsFile();

    // 文件格式
    void exportToCsv();
    void importFromCsv();

    void about();

    void onCurrentWorksheetChanged(int index);

private:
    void setupUi();
    void setupMenus();
    void setupToolbars();
    void connectSignals();
    bool maybeSave();
    void setCurrentFile(const QString &fileName);

    WorksheetManager *m_worksheetManager;
    SearchWidget *m_searchWidget;
    std::unique_ptr<Workbook> m_workbook;

    QString m_currentFileName; // 当前文件的路径
    bool m_isModified; // 修改标志
};
