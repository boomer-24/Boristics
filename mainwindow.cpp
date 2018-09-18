#include "mainwindow.h"
#include "ui_mainwindow.h"

MainWindow::MainWindow(QWidget *parent) : QMainWindow(parent), ui_(new Ui::MainWindow)
{
    ui_->setupUi(this);
//    MoveHandler MH;
//    MH.Initialize("ini.xml");
//    MH.TraverseArchiveDir("C:\\Users\\user\\Desktop\\BORIS\\Formula-2K\\ОФОРМИТЬ И В АРХИВ\\2K");
    ExcelHandler EH("C:\\Users\\user\\Desktop\\BORIS\\БД F2K.xls");
    QStringList sl = EH.getNewProgram(QDate(2015, 01, 10));
    qDebug() << sl;
}

MainWindow::~MainWindow()
{
    delete ui_;
}

void MainWindow:: ProcessOnePrgDir(const QString &_prgDirPath)
{

    /*СНАЧАЛА НАЙТИ *.DOCX, УЗНАТЬ 1)ТЕСТЕРЫ, 2)СЕРИЮ, 3)НАЗВАНИЕ */
    /*ТУПО СКОРМИТЬ SRC-ПАПКУ И DST-ПАПКУ moveHandler`у*/

}

