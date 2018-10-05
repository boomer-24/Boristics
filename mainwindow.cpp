#include "mainwindow.h"
#include "ui_mainwindow.h"

MainWindow::MainWindow(QWidget *parent) : QMainWindow(parent), ui_(new Ui::MainWindow)
{
    ui_->setupUi(this);
    this->close();
    this->setWindowTitle("Boristik");
    this->trayIcon_ = new QSystemTrayIcon(this);
    this->trayIcon_->setIcon(this->style()->standardIcon(QStyle::SP_TrashIcon));
    this->trayIcon_->setToolTip("Tray Program \n Работа со сворачиванием программы трей");
    QMenu* menu = new QMenu(this);
    QAction* viewWindow = new QAction(trUtf8("Развернуть окно"), this);
    QAction* quitAction = new QAction(trUtf8("Выход"), this);
    QObject::connect(viewWindow, SIGNAL(triggered()), this, SLOT(show()));
    QObject::connect(quitAction, SIGNAL(triggered()), this, SLOT(close()));

    menu->addAction(viewWindow);
    menu->addAction(quitAction);

    this->trayIcon_->setContextMenu(menu);
    this->trayIcon_->show();

    QObject::connect(this->trayIcon_, SIGNAL(activated(QSystemTrayIcon::ActivationReason)),
                     this, SLOT(iconActivated(QSystemTrayIcon::ActivationReason)));

    this->Initialize(QCoreApplication::applicationDirPath().append("/ini.xml"));
    this->AutoRun(this->autorun_);

    this->thread_ = new QThread(this);
    this->MH_ = new MoveHandler();
    QObject::connect(this, SIGNAL(signalStartOperation()), this->MH_, SLOT(slotStartOperations()));
    QObject::connect(this->MH_, SIGNAL(signalInfoToUItrueTextBox(QString)), this, SLOT(slotInfoToUItrueTextBox(QString)));
    QObject::connect(this->MH_, SIGNAL(signalInfoToUIfailTextBox(QString)), this, SLOT(slotInfoToUIfailTextBox(QString)));
    QObject::connect(this->MH_, SIGNAL(signalTraverseArchiveComplete()), this->ui_->pushButton_go, SLOT(show()));
    QObject::connect(this->MH_, SIGNAL(signalProgressToUI(int)), this->ui_->progressBar, SLOT(setValue(int)));
    this->thread_->start();
    this->MH_->moveToThread(this->thread_);

    this->threadNewPrograms_ = new QThread(this);
    this->FNP_ = new FinderNewPrograms();
    QObject::connect(this, SIGNAL(signalStartOperationFindNewProgram(QString,QDate)),
                     this->FNP_, SLOT(slotStartOperation(QString,QDate)));
    QObject::connect(this->FNP_, SIGNAL(signalInfoToUInewProgramTrueTextBox(QString)),
                     this, SLOT(slotAppendToNewProgram(QString)));
    QObject::connect(this->FNP_, SIGNAL(signalInfoToUInewProgramFailTextBox(QString)),
                     this->ui_->textBrowser_fails, SLOT(append(QString)));
    QObject::connect(this->FNP_, SIGNAL(signalCurrentSheetToUI(QString)),
                     this, SLOT(slotObtainCurrentSheetForUI(QString)));
    QObject::connect(this->FNP_, SIGNAL(signalSearchNewProgramComplete()),
                     this->ui_->pushButton_getNewPrograms, SLOT(show()));
    QObject::connect(this->FNP_, SIGNAL(signalSearchNewProgramComplete()),
                     this->ui_->label_currentSeries, SLOT(clear()));

    this->FNP_->moveToThread(this->threadNewPrograms_);
    this->threadNewPrograms_->start();
}

MainWindow::~MainWindow()
{
    delete this->MH_;
    this->thread_->quit();
    this->threadNewPrograms_->quit();
    //        delete this->thread_; //НЕ ПОНЯЛ ПОКА ПОЧЕМУ, НО ТАК ДЕЛАТЬ НЕ СТОИТ
    delete ui_;
}

void MainWindow::Initialize(const QString &_xmlPath)
{
    QDomDocument domDoc;
    QFile file(_xmlPath);
    if (file.open(QIODevice::ReadOnly))
    {
        if (domDoc.setContent(&file))
        {
            QDomElement domElement = domDoc.documentElement();
            QDomNode domNode = domElement.firstChild();
            while(!domNode.isNull())
            {
                if (domNode.isElement())
                {
                    QDomElement domElement = domNode.toElement();
                    if (!domElement.isNull())
                    {
                        if (domElement.tagName() == "from2K")
                        {
                            this->path2Kfrom_ = domElement.text();
                        }
                        else if (domElement.tagName() == "todocs2K")
                        {
                            this->path2Kdocs_ = domElement.text();
                        }
                        else if (domElement.tagName() == "toprogs2K")
                        {
                            this->path2Kprgs_ = domElement.text();
                        }
                        else if (domElement.tagName() == "excel2K")
                        {
                            this->path2Kexcel_ = domElement.text();
                        }
                        else if (domElement.tagName() == "autorun")
                        {
                            QString autorunStr(domElement.text());
                            this->autorun_ = autorunStr.toInt();
                        }
                        else qDebug() << "Tag ne found";
                    }
                    domNode = domNode.nextSibling();
                }
            }
        } else this->slotInfoToUIfailTextBox("It`s no XML!");
    } else this->slotInfoToUIfailTextBox("File not open =/");
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
    if (!this->MH_->isExcelBusy())
    {
        QString AppPath(QCoreApplication::applicationDirPath());
        AppPath.append("/tasklist.txt");
        QString query("tasklist > ");
        query.append(AppPath);
        system(query.toStdString().data());
        QFile file(AppPath);
        if (file.open(QIODevice::ReadOnly))
        {
            QByteArray data;
            data = file.readAll();
            QString str(data);
            str = str.toLower();
            if (str.contains("excel"))
            {
                QMessageBox::StandardButton reply = QMessageBox::warning(this, "\\m/", "Для дальнейшей работы закрой програму Excel!");
                return;
            }
        } else
        {
            QMessageBox::StandardButton reply = QMessageBox::warning(this, "\\m/", "tasklist не считан. Обратись к создателю.");
            return;
        }
        if (this->ui_->calendarWidget->selectedDate() > QDate::currentDate())
        {
            QMessageBox::StandardButton reply = QMessageBox::warning(this, "\\m/", "Выбрана дата будущего. Странно.");
            return;
        }
        if (this->ui_->calendarWidget->selectedDate() == QDate::currentDate())
        {
            QMessageBox::StandardButton reply = QMessageBox::question(this, "\\m/", "Поиск будет с сегодняшней даты. Продолжить?",
                                                                      QMessageBox::Yes | QMessageBox::No);
            if (reply == QMessageBox::No)
                return;
        }
        this->ui_->pushButton_getNewPrograms->hide();
        emit this->signalStartOperationFindNewProgram("C:/Users/user/Desktop/BORIS/БД F2K.xls",
                                                      this->ui_->calendarWidget->selectedDate());

    } else
        QMessageBox::StandardButton reply = QMessageBox::warning(this, "\\m/", "В таблицу сейчас производится запись. "
                                                                               "Дождись пока закончится и попробуй еще.");

}

void MainWindow::on_pushButton_go_clicked()
{
    if (!this->FNP_->isExcelBusy())
    {
        QString AppPath(QCoreApplication::applicationDirPath());
        AppPath.append("/tasklist.txt");
        QString query("tasklist > ");
        query.append(AppPath);
        system(query.toStdString().data());
        QFile file(AppPath);
        if (file.open(QIODevice::ReadOnly))
        {
            QByteArray data;
            data = file.readAll();
            QString str(data);
            str = str.toLower();
            if (str.contains("excel") && str.contains("word"))
            {
                QMessageBox::StandardButton reply = QMessageBox::warning(this, "\\m/", "Для дальнейшей работы закрой програмы Excel и Word!");
                return;
            }
            if (str.contains("excel"))
            {
                QMessageBox::StandardButton reply = QMessageBox::warning(this, "\\m/", "Для дальнейшей работы закрой програму Excel!");
                return;
            }
            if (str.contains("word"))
            {
                QMessageBox::StandardButton reply = QMessageBox::warning(this, "\\m/", "Для дальнейшей работы закрой програму Word!");
                return;
            }
        } else
        {
            QMessageBox::StandardButton reply = QMessageBox::warning(this, "\\m/", "tasklist не считан. Обратись к создателю.");
            return;
        }

        this->ui_->pushButton_go->hide();
        emit this->signalStartOperation();
    } else
        QMessageBox::StandardButton reply = QMessageBox::warning(this, "\\m/", "Сейчас производится поиск новых программ. "
                                                                               "Дождись пока закончится и попробуй еще.");
}

void MainWindow::on_pushButton_clearNewProgramsTextBox_clicked()
{
    this->ui_->textBrowser_newPrograms->clear();
    this->ui_->label_amountNewElements->clear();
    this->ui_->label_amountNewElementsText->clear();
}

void MainWindow::slotInfoToUItrueTextBox(const QString &_info)
{    
    this->ui_->textBrowser_success->append(_info);
}

void MainWindow::slotInfoToUIfailTextBox(const QString &_infoError)
{
    this->ui_->textBrowser_complaints->append(_infoError);
}

void MainWindow::slotAppendToNewProgram(const QString &_infoNewProgram)
{
    this->ui_->textBrowser_newPrograms->append(_infoNewProgram);
    int amount = this->ui_->label_amountNewElements->text().toInt();
    amount++;
    this->ui_->label_amountNewElements->setNum(amount);
    QString amountStr(QString::number(amount));
    QString element("элемент");
    QString lastCharacter(amountStr.at(amountStr.size() - 1));
    if (lastCharacter.contains(QRegExp("([5-9]|0){1,1}")))
        element.append("ов");
    else if (lastCharacter.contains(QRegExp("[2-4]{1,1}")))
        element.append("а");
    this->ui_->label_amountNewElementsText->setText(element);
}

void MainWindow::slotAppendToNewProgramFail(const QString &_infoNewProgramFail)
{
    this->ui_->textBrowser_fails->append(_infoNewProgramFail);
}

void MainWindow::slotObtainCurrentSheetForUI(const QString &_element)
{
    this->ui_->label_currentSeries->setText(_element);
}

void MainWindow::AutoRun(bool _isAutorun)
{
    if (_isAutorun)
    {
        QSettings setting("HKEY_CURRENT_USER\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run", QSettings::NativeFormat);
        setting.setValue("BoristikApp", QDir::toNativeSeparators(QApplication::applicationFilePath()));
        setting.sync();
    } else
    {
        QSettings setting("HKEY_CURRENT_USER\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run", QSettings::NativeFormat);
        setting.setValue("BoristikApp", QDir::toNativeSeparators(QApplication::applicationFilePath()));
        setting.remove("BoristikApp");
    }
}
