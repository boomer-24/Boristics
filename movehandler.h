#ifndef MOVEHANDLER_H
#define MOVEHANDLER_H

#include "pasport2Kanalyzer.h"
#include "excelhandler.h"

#include <QString>
#include <QDir>
#include <QFile>
#include <QtXml>

class MoveHandler
{
private:
//    QString pathFrom_, pathTo_, pathBase_, elementName_;
    QString path2Kfrom_, path2Kdocs_, path2Kprgs_, path2Kexcel_;

public:
    MoveHandler();    
    ~MoveHandler();
    void Initialize(const QString &_xmlPath);
    void TraverseArchiveDir(const QString &_dirPath);
    void TraverseMainProgramDir(const QString &_dirPath);
    const QString FindProgramPasportPath(const QString &_dirPath);
    void DirsCopy(const QString &_dirPathFrom, const QString &_dirPathTo);
    void setPath2Kdocs(const QString &_path2Kdocs);
    void setPath2Kprgs(const QString &_path2Kprgs);
    void setPath2Kfrom(const QString &_path2Kfrom);
};

#endif // MOVEHANDLER_H
