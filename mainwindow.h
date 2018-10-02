#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include "Passport2kanalyzer.h"
#include "excelhandler.h"
#include "movehandler.h"
#include <QMainWindow>
#include <QFileDialog>
#include <QCloseEvent>
#include <QSystemTrayIcon>
#include <QAction>
#include <QMessageBox>

namespace Ui {
class MainWindow;
}

class MainWindow : public QMainWindow
{
    Q_OBJECT

public:
    explicit MainWindow(QWidget *parent = 0);
    ~MainWindow();
private slots:
    void iconActivated(QSystemTrayIcon::ActivationReason _reason);
    void on_pushButton_ok_clicked();
    void on_pushButton_go_clicked();
    void on_pushButton_getNewPrograms_clicked();
public slots:
    void slotInfoToUItrueTextBox(QString info);
    void slotInfoToUIfailTextBox(QString infoError);
signals:
    void signalStartOperation();

protected:
    void closeEvent(QCloseEvent *event) override;
    void resizeEvent(QResizeEvent *event) override;
private:
    Ui::MainWindow *ui_;
    QSystemTrayIcon *trayIcon_;
    MoveHandler *MH_;
    QThread *thread_;
    QThread *threadNewPrograms_;
};

#endif // MAINWINDOW_H
