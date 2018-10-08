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
#include <QTimer>

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
    void AutoRun(bool _isAutorun) const;
    void WriteXML(const QString &_xmlPath);

private slots:
    void iconActivated(QSystemTrayIcon::ActivationReason _reason);    
    void on_pushButton_go_clicked();
    void on_pushButton_getNewPrograms_clicked();
    void on_pushButton_clearNewProgramsTextBox_clicked();
    void on_pushButton_clearNewProgramsFailTextBox_clicked();
    void on_pushButton_savePaths_clicked();

    void on_pushButton_ok_clicked();

public slots:
    void slotInfoToUItrueTextBox(const QString& _info);
    void slotInfoToUIfailTextBox(const QString& _infoError);
    void slotAppendToNewProgram(const QString& _infoNewProgram);
    void slotAppendToNewProgramFail(const QString& _infoNewProgramFail);
    void slotObtainCurrentSheetForUI(const QString& _element);
    void slotBalloonClickedGetNewPtogramShow();
    void slotShowNoticeAboutNewProgram();

signals:
    void signalStartOperation();    
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
    bool autorun_, needForReminders_;
    QTimer timer_;
    int reminderFrequencyMinutes_, dayBoard_;
};

#endif // MAINWINDOW_H
