#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include "Passport2kanalyzer.h"
#include "excelhandler.h"
#include "movehandler.h"
#include "findernewprograms.h"
#include <QMainWindow>
#include <QFileDialog>
#include <QCloseEvent>
#include <QSystemTrayIcon>
#include <QAction>
#include <QMessageBox>
#include <thread>

namespace Ui {
class MainWindow;
}

class MainWindow : public QMainWindow
{
    Q_OBJECT

public:
    explicit MainWindow(QWidget *parent = 0);
    ~MainWindow();
    void Initialize(const QString &_xmlPath);

private slots:
    void iconActivated(QSystemTrayIcon::ActivationReason _reason);
    void on_pushButton_ok_clicked();
    void on_pushButton_go_clicked();
    void on_pushButton_getNewPrograms_clicked();

    void on_pushButton_clearNewProgramsTextBox_clicked();

public slots:
    void slotInfoToUItrueTextBox(QString info);
    void slotInfoToUIfailTextBox(QString infoError);
    void slotAppendToNewProgram(QString infoNewProgram);
    void slotAppendToNewProgramFail(QString _infoNewProgramFail);
    void slotObtainCurrentSheetForUI(QString _element);

signals:
    void signalStartOperation();
    void signalStartOperationFindNewProgram();
    void signalStartOperationFindNewProgram(QString,QDate);

protected:
    void closeEvent(QCloseEvent *event) override;
    void resizeEvent(QResizeEvent *event) override;

private:
    Ui::MainWindow *ui_;
    QSystemTrayIcon *trayIcon_;
    MoveHandler *MH_;
    FinderNewPrograms *FNP_;
    QThread *thread_;
    QThread *threadNewPrograms_;
    QString path2Kfrom_, path2Kdocs_, path2Kprgs_, path2Kexcel_;
};

#endif // MAINWINDOW_H
