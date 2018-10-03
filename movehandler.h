#ifndef MOVEHANDLER_H
#define MOVEHANDLER_H

#include "passport2Kanalyzer.h"
#include "excelhandler.h"

#include <QObject>
#include <QString>
#include <QDir>
#include <QFile>
#include <QtXml>

class MoveHandler : public QObject
{
    Q_OBJECT
private:
    QString path2Kfrom_, path2Kdocs_, path2Kprgs_, path2Kexcel_;
    bool isExcelBusy_;

public:
    explicit MoveHandler(QObject *parent = 0);
    ~MoveHandler();

    void Initialize(const QString &_xmlPath);
    void TraverseArchiveDir();
    void TraverseArchiveDir(const QString &_dirPath);
    bool TraverseMainProgramDir(const QString &_dirPath);
    const QString FindProgramPassportPath(const QString &_dirPath);
    void DirsCopy(const QString &_dirPathFrom, const QString &_dirPathTo);
    void PassportCopy(const QString &_pathFrom, const QString &_series);
    void setPath2Kdocs(const QString &_path2Kdocs);
    void setPath2Kprgs(const QString &_path2Kprgs);
    void setPath2Kfrom(const QString &_path2Kfrom);
    QString path2Kexcel() const;
    bool isExcelBusy() const;

public slots:
    void slotStartOperations();

signals:
    void signalInfoToUItrueTextBox(QString info);
    void signalToUIfailTextBox(QString infoErr);
    void signalTraverseArchiveComplete();
    void signalStart();
    void signalFinish();
    void signalProgressToUI(int);
};

#endif // MOVEHANDLER_H
