#include <QApplication>
#include "ui/MainWindow.h"

int main(int argc, char *argv[])
{
    QApplication app(argc, argv);

    app.setApplicationName("Spreadsheet"); // 设置应用程序进程名称
    app.setApplicationVersion("1.0.0");
    
    MainWindow w;
    w.show();

    return app.exec();
}
