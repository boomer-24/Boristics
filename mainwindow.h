#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include "pasport2kanalyzer.h"
#include "excelhandler.h"
#include "movehandler.h"
#include <QMainWindow>
#include <QFileDialog>

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

private:
    Ui::MainWindow *ui_;
};

#endif // MAINWINDOW_H
