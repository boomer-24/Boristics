#include "mainwindow.h"
#include "ui_mainwindow.h"

MainWindow::MainWindow(QWidget *parent) : QMainWindow(parent), ui_(new Ui::MainWindow)
{
    ui_->setupUi(this);
//    Pasport2KAnalyzer pasportAnalyzer;
////    //    this->pasportAnalyzer_.SetDocxPath("C:/Users/user/Desktop/boristika/1821RE55-0015_ТУ.docx");
////    //    this->pasportAnalyzer_.SetDocxPath("C:/Users/user/Desktop/boristika/1821RE55-0015_ТУ.docx");
//    pasportAnalyzer.SetDocxPath(QFileDialog::getOpenFileName(this, "Select file", "C:/", "*.doc *.docx"));
//    if (pasportAnalyzer.DocumentsCount())
//    {
//        pasportAnalyzer.Initialize();
//        //    ExcelHandler EH();
//        QString path("C:/Users/user/Desktop/boristika/БД F2K.xls");
//        ExcelHandler EH(path);
//        const auto a = pasportAnalyzer.rowsInExcelTable();
//        for (int i = 0; i < a.size(); i++)
//        {
//            EH.InsertRow(a.at(i));
//        }
//    } else qDebug() << "Documents______:" << pasportAnalyzer.DocumentsCount();
//    QString path(QFileDialog::getOpenFileName(this, "Select file", "C:/", "*.doc *.docx"));
//    QString path(QFileDialog::getExistingDirectory(this, "Set dir"));
    MoveHandler MH;
    MH.Initialize("ini.xml");
    MH.TraverseArchiveDir("C:\\Users\\user\\Desktop\\BORIS\\Formula-2K\\ОФОРМИТЬ И В АРХИВ\\2K");

//    QString pathFrom(QFileDialog::getExistingDirectory(this, "FROM dir"));
//    QString pathTo(QFileDialog::getExistingDirectory(this, "TO dir"));
//    MH.DirsCopy(pathFrom, pathTo);
}

MainWindow::~MainWindow()
{
    delete ui_;
}

void MainWindow:: ProcessOnePrgDir(const QString &_prgDirPath)
{

    /*СНАЧАЛА НАЙТИ *.DOCX, УЗНАТЬ 1)ТЕСТЕРЫ, 2)СЕРИЮ, 3)НАЗВАНИЕ */
    /*ТУПО СКОРМИТЬ SRC-ПАПКУ И DST-ПАПКУ moveHandlerу*/

}

void MainWindow::on_pushButton_move_clicked()
{

}
