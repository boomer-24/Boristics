#include "mainwindow.h"
#include "ui_mainwindow.h"

MainWindow::MainWindow(QWidget *parent) : QMainWindow(parent), ui_(new Ui::MainWindow)
{
    ui_->setupUi(this);
    this->setWindowTitle("Boristik");
    this->trayIcon_ = new QSystemTrayIcon(this);
    this->trayIcon_->setIcon(this->style()->standardIcon(QStyle::SP_TrashIcon));
    this->trayIcon_->setToolTip("Tray Program \n Работа со сворачиванием программы трей");
    QMenu * menu = new QMenu(this);
    QAction * viewWindow = new QAction(trUtf8("Развернуть окно"), this);
    QAction * quitAction = new QAction(trUtf8("Выход"), this);
    connect(viewWindow, SIGNAL(triggered()), this, SLOT(show()));
    connect(quitAction, SIGNAL(triggered()), this, SLOT(close()));

    menu->addAction(viewWindow);
    menu->addAction(quitAction);

    this->trayIcon_->setContextMenu(menu);
    this->trayIcon_->show();

    connect(this->trayIcon_, SIGNAL(activated(QSystemTrayIcon::ActivationReason)),
            this, SLOT(iconActivated(QSystemTrayIcon::ActivationReason)));        

    this->thread_ = new QThread(this);
    this->MH_ = new MoveHandler();
    QObject::connect(this, SIGNAL(signalStartOperation()), this->MH_, SLOT(slotStartOperations()));
    QObject::connect(this, SIGNAL(signalStartOperation()), this->ui_->pushButton_go, SLOT(hide()));
    QObject::connect(this->MH_, SIGNAL(signalFinish()), this->ui_->pushButton_go, SLOT(show()));
    QObject::connect(this, SIGNAL(destroyed(QObject*)), this->thread_, SLOT(quit()));
    QObject::connect(this->MH_, SIGNAL(signalInfoToUItrueTextBox(QString)), this, SLOT(slotInfoToUItrueTextBox(QString)));
    QObject::connect(this->MH_, SIGNAL(signalToUIfailTextBox(QString)), this, SLOT(slotInfoToUIfailTextBox(QString)));
    QObject::connect(this->MH_, SIGNAL(signalStart()), this->ui_->pushButton_go, SLOT(hide()));
    QObject::connect(this->MH_, SIGNAL(signalStart()), this->thread_, SLOT(start()));
    QObject::connect(this->MH_, SIGNAL(signalFinish()), this->ui_->pushButton_go, SLOT(show()));
    QObject::connect(this->MH_, SIGNAL(signalFinish()), this->thread_, SLOT(quit()));
    QObject::connect(this->MH_, SIGNAL(signalTraverseArchiveComplete()), this->ui_->pushButton_go, SLOT(show()));
    QObject::connect(this->MH_, SIGNAL(signalProgressToUI(int)), this->ui_->progressBar, SLOT(setValue(int)));
    this->MH_->moveToThread(this->thread_);
    this->thread_->start();
}

MainWindow::~MainWindow()
{
    delete this->MH_;
//    delete this->thread_;
    delete ui_;
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

void MainWindow::closeEvent(QCloseEvent *event)
{
    if (this->isVisible())
    {
        event->ignore();
        this->hide();
        QSystemTrayIcon::MessageIcon icon = QSystemTrayIcon::MessageIcon(QSystemTrayIcon::Information);

        this->trayIcon_->showMessage("Tray Program", trUtf8("Приложение свернуто в трей."), icon, 200);
    }
}

void MainWindow::resizeEvent(QResizeEvent *event)
{
    this->ui_->tabWidget->setFixedSize(this->size() - QSize(20, 50));
    this->ui_->textBrowser_success->setFixedSize(this->ui_->tabWidget->size() - QSize(10 + this->ui_->textBrowser_success->x(),
                                                                              10 + this->ui_->textBrowser_success->y() / 2 +
                                                                                      this->ui_->tabWidget->height() / 2 + 40));
    this->ui_->textBrowser_complaints->move(this->ui_->textBrowser_complaints->x(), this->ui_->tabWidget->height() / 2);
    this->ui_->textBrowser_complaints->setFixedSize(this->ui_->textBrowser_success->size());

    this->ui_->textBrowser_newPrograms->setFixedSize(this->ui_->tabWidget->size() -
                                                     QSize(this->ui_->textBrowser_newPrograms->pos().rx(),
                                                           this->ui_->textBrowser_newPrograms->pos().ry()) -
                                                     QSize(10, 30));
    this->ui_->textBrowser_fails->setFixedHeight(this->ui_->tabWidget->height() - this->ui_->textBrowser_fails->pos().ry() - 30);
}

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

void MainWindow::on_pushButton_getNewPrograms_clicked()
{    
    if (this->ui_->calendarWidget->selectedDate() == QDate::currentDate())
    {
        QMessageBox::StandardButton reply = QMessageBox::question(this, "\\m/", "Поиск будет с сегодняшней даты. Продолжить?",
                                                                  QMessageBox::Yes | QMessageBox::No);
        if (reply == QMessageBox::No)
            return;
    }
    this->ui_->textBrowser_newPrograms->clear();
    this->threadNewPrograms_ = new QThread(this);

    ExcelHandler EH("C:\\Users\\user\\Desktop\\BORIS\\БД F2K.xls");
    QStringList sl = EH.getNewProgram(this->ui_->calendarWidget->selectedDate());
    QString element("элемент");
    QString amount(QString::number(sl.size()));
    QString lastCharacter(amount.at(amount.size() - 1));
    if (lastCharacter.contains(QRegExp("([5-9]|0){1,1}")))
        element.append("ов");
    else if (lastCharacter.contains(QRegExp("[2-4]{1,1}")))
        element.append("а");
    this->ui_->label_amountNewElements->setText(amount.append(" ").append(element));
    for (const QString& str : sl)
        this->ui_->textBrowser_newPrograms->append(str);
}

void MainWindow::slotInfoToUItrueTextBox(QString info)
{
    this->ui_->textBrowser_success->append(info);
}

void MainWindow::slotInfoToUIfailTextBox(QString infoError)
{
    this->ui_->textBrowser_complaints->append(infoError);
}

void MainWindow::on_pushButton_go_clicked()
{
//    this->MH_ = new MoveHandler();
//    this->thread_ = new QThread(this);
//    QObject::connect(this, SIGNAL(destroyed(QObject*)), this->thread_, SLOT(quit()));
//    QObject::connect(this->MH_, SIGNAL(signalInfoToUItrueTextBox(QString)), this, SLOT(slotInfoToUItrueTextBox(QString)));
//    QObject::connect(this->MH_, SIGNAL(signalStart()), this->ui_->pushButton_go, SLOT(hide()));
//    QObject::connect(this->MH_, SIGNAL(signalFinish()), this->ui_->pushButton_go, SLOT(show()));
//    QObject::connect(this->MH_, SIGNAL(signalFinish()), this->thread_, SLOT(quit()));
//    qDebug() << "before " << this->MH_->thread();
//    this->MH_->moveToThread(this->thread_);
//    qDebug() << "after " << this->MH_->thread();
//    this->thread_->start();
//    this->MH_->Initialize(QCoreApplication::applicationDirPath().append("/ini.xml"));
//    this->MH_->TraverseArchiveDir();
//    this->ui_->pushButton_go->hide();
    emit this->signalStartOperation();
}
