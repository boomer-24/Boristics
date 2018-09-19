#include "mainwindow.h"
#include "ui_mainwindow.h"

MainWindow::MainWindow(QWidget *parent) : QMainWindow(parent), ui_(new Ui::MainWindow)
{
    ui_->setupUi(this);
//    this->trayIcon_ = new QSystemTrayIcon(this);
//    this->trayIcon_->setIcon(this->style()->standardIcon(QStyle::SP_ComputerIcon));
//    this->trayIcon_->setToolTip("Tray Program \n Работа со сворачиванием программы трей");
//    QMenu * menu = new QMenu(this);
//    QAction * viewWindow = new QAction(trUtf8("Развернуть окно"), this);
//    QAction * quitAction = new QAction(trUtf8("Выход"), this);
//    connect(viewWindow, SIGNAL(triggered()), this, SLOT(show()));
//    connect(quitAction, SIGNAL(triggered()), this, SLOT(close()));

//    menu->addAction(viewWindow);
//    menu->addAction(quitAction);

//    this->trayIcon_->setContextMenu(menu);
//    this->trayIcon_->show();

//    connect(this->trayIcon_, SIGNAL(activated(QSystemTrayIcon::ActivationReason)),
//            this, SLOT(iconActivated(QSystemTrayIcon::ActivationReason)));
//    connect(this->trayIcon_, SIGNAL(messageClicked()), this->trayIcon_, SLOT(hide()));

        MoveHandler MH;
        MH.Initialize("ini.xml");
        MH.TraverseArchiveDir("C:\\Users\\user\\Desktop\\BORIS\\Formula-2K\\ОФОРМИТЬ И В АРХИВ\\2K");

    //    ExcelHandler EH("C:\\Users\\user\\Desktop\\BORIS\\БД F2K.xls");
    //    QStringList sl = EH.getNewProgram(QDate(2015, 01, 10));
    //    qDebug() << sl;
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

void MainWindow::iconActivated(QSystemTrayIcon::ActivationReason _reason)
{
    switch (_reason)
    {
    case QSystemTrayIcon::Trigger:
        if (!this->isVisible())
            this->show();
        else
            this->hide();
        break;
    case QSystemTrayIcon::DoubleClick:
    {
        QSystemTrayIcon::MessageIcon icon = QSystemTrayIcon::MessageIcon(QSystemTrayIcon::Critical);
        this->trayIcon_->showMessage("Tray Program", trUtf8("ДАБЛ КЛИК!!!!"), icon, 200);
    }
    default:
        break;
    }
}

//void MainWindow::closeEvent(QCloseEvent *event)
//{
//    if (this->isVisible())
//    {
//        event->ignore();
//        this->hide();
//        QSystemTrayIcon::MessageIcon icon = QSystemTrayIcon::MessageIcon(QSystemTrayIcon::Information);

//        this->trayIcon_->showMessage("Tray Program", trUtf8("Приложение свернуто в трей."), icon, 200);
//    }
//}

void MainWindow::on_pushButton_ok_clicked()
{
    QSystemTrayIcon::MessageIcon icon = QSystemTrayIcon::MessageIcon(QSystemTrayIcon::Warning);
    this->trayIcon_->showMessage("Boristic", trUtf8("Нажми на меня и напишем служебку в 12 отделение!\nДАВАЙ, ДАВАЙ, ДАВАЙ, ДАВАЙ, ДАВАЙ, ДАВАЙ, ДАВАЙ БРЮ"
                                                    "Нажми на меня и напишем служебку в 12 отделение!\nДАВАЙ, ДАВАЙ, ДАВАЙ, ДАВАЙ, ДАВАЙ, ДАВАЙ, ДАВАЙ БРЮ"
                                                    "Нажми на меня и напишем служебку в 12 отделение!\nДАВАЙ, ДАВАЙ, ДАВАЙ, ДАВАЙ, ДАВАЙ, ДАВАЙ, ДАВАЙ БРЮ"
                                                    "Нажми на меня и напишем служебку в 12 отделение!\nДАВАЙ, ДАВАЙ, ДАВАЙ, ДАВАЙ, ДАВАЙ, ДАВАЙ, ДАВАЙ БРЮ"
                                                    "Нажми на меня и напишем служебку в 12 отделение!\nДАВАЙ, ДАВАЙ, ДАВАЙ, ДАВАЙ, ДАВАЙ, ДАВАЙ, ДАВАЙ БРЮ"
                                                    "Нажми на меня и напишем служебку в 12 отделение!\nДАВАЙ, ДАВАЙ, ДАВАЙ, ДАВАЙ, ДАВАЙ, ДАВАЙ, ДАВАЙ БРЮ"
                                                    "Нажми на меня и напишем служебку в 12 отделение!\nДАВАЙ, ДАВАЙ, ДАВАЙ, ДАВАЙ, ДАВАЙ, ДАВАЙ, ДАВАЙ БРЮ"
                                                    "Нажми на меня и напишем служебку в 12 отделение!\nДАВАЙ, ДАВАЙ, ДАВАЙ, ДАВАЙ, ДАВАЙ, ДАВАЙ, ДАВАЙ БРЮ"
                                                    "Нажми на меня и напишем служебку в 12 отделение!\nДАВАЙ, ДАВАЙ, ДАВАЙ, ДАВАЙ, ДАВАЙ, ДАВАЙ, ДАВАЙ БРЮ"
                                                    "Нажми на меня и напишем служебку в 12 отделение!\nДАВАЙ, ДАВАЙ, ДАВАЙ, ДАВАЙ, ДАВАЙ, ДАВАЙ, ДАВАЙ БРЮ!"), icon, 500);
}
