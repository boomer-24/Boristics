#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include "pasport2kanalyzer.h"
#include "excelhandler.h"
#include "movehandler.h"
#include <QMainWindow>
#include <QFileDialog>
#include <QCloseEvent>
#include <QSystemTrayIcon>
#include <QAction>

namespace Ui {
class MainWindow;
}

class MainWindow : public QMainWindow
{
    Q_OBJECT

public:
    explicit MainWindow(QWidget *parent = 0);
    ~MainWindow();    
    void ProcessOnePrgDir(const QString& _prgDirPath);  // тоже нах
private slots:
    void iconActivated(QSystemTrayIcon::ActivationReason _reason);
    void on_pushButton_ok_clicked();

//protected:
//    void closeEvent(QCloseEvent *event) override;

private:
    Ui::MainWindow *ui_;
    QSystemTrayIcon *trayIcon_;
};

#endif // MAINWINDOW_H
