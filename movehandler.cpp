#include "movehandler.h"

MoveHandler::MoveHandler()
{

}

MoveHandler::~MoveHandler()
{

}

void MoveHandler::Initialize(const QString &_xmlPath)
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
                        else qDebug() << "Tag ne found";
                    }
                    domNode = domNode.nextSibling();
                }
            }
        } else qDebug() << "It`s no XML!";
    } else qDebug() << "File not open =/";
}

void MoveHandler::TraverseArchiveDir(const QString &_dirPath) //    ПЕРЕДАВАТЬ ПАПКУ ОФОРМИТЬ И В АРХИВ/2К
{
    const QDir dir(_dirPath);
    if (!dir.exists())
        qDebug() << "No such directory : %s", dir.dirName();
    else
    {
        QStringList slEntryDir = dir.entryList(QDir::Dirs);
        slEntryDir.removeOne(".");
        slEntryDir.removeOne("..");
        for (const QString& dirName : slEntryDir)
        {
            QString dirPath = dir.absolutePath();
            QString dirEntryPath = dirPath.append("/").append(dirName);
            this->TraverseMainProgramDir(dirEntryPath); // ДЛЯ БАЗОВОЙ ПАПКИ ОДНОЙ ПРОГРАММЫ
        }
    }
}

void MoveHandler::TraverseMainProgramDir(const QString &_dirPath) // ДЛЯ БАЗОВОЙ ПАПКИ ОДНОЙ ПРОГРАММЫ
{
    const QDir dir(_dirPath);
    if (!dir.exists())
        qDebug() << "No such directory : %s", dir.dirName();
    else
    {
        QString docxPath(this->FindProgramPasportPath(_dirPath));
        if (!docxPath.isEmpty())
        {
            QStringList slPath(docxPath.split("."));
            QFile docxFile(docxPath);
            QString newPath(QCoreApplication::applicationDirPath().append("/pasport").append(".").append(slPath.last()));
            if (QFile::exists(newPath))
                QFile::remove(newPath);

            docxFile.copy(newPath);
            Pasport2KAnalyzer pasport(newPath);
            if (pasport.DocumentsCount())
            {
                pasport.Initialize();
                ExcelHandler excelHandler(this->path2Kexcel_);
                const QVector<RowInExcelTable> &vRows = pasport.rowsInExcelTable();
                pasport.QuitWord();
                if (QFile::exists(newPath))
                    QFile::remove(newPath);
                QString dirNameForMoving(_dirPath.split("/").last());
                if (!dirNameForMoving.contains(QRegExp("\\d+\\s\\w")))
                {
                    QRegExp rgxp("(\\d+)(\\w+)");
                    int pos = rgxp.indexIn(dirNameForMoving);
                    dirNameForMoving.insert(rgxp.pos(2), " ");
                }
                for (RowInExcelTable row : vRows)
                {
                    //    excelHandler.InsertRow(row);
                    const QString &series(row.series());
                    const QStringList &slTestersFromPasport(row.slTesters());
                    const QStringList &slTestersExistingDirs(QDir(this->path2Kprgs_).entryList(QDir::Dirs));//nodotanddotdot

                    for (const QString &testerExistingDirs : slTestersExistingDirs)
                        for (const QString &testerFromPasport : slTestersFromPasport)
                        {
                            if (testerExistingDirs.contains(testerFromPasport))
                            {
                                QString progsArchiveOneTester(this->path2Kprgs_);
                                progsArchiveOneTester.append("/").append(testerExistingDirs);
                                QDir dir(progsArchiveOneTester);
                                QStringList slSeries(dir.entryList(QDir::Dirs));
                                if (!slSeries.contains(series))
                                    dir.mkdir(series);
                                QDir dirInsideOneSeries(progsArchiveOneTester + ("/") + series);
                                QStringList slProgramsDir(dirInsideOneSeries.entryList(QDir::Dirs));
                                while (slProgramsDir.contains(dirNameForMoving))
                                    dirNameForMoving.append("_Newer");
                                QString finishPath(dirInsideOneSeries.absolutePath().append("/").append(dirNameForMoving));
                                dirInsideOneSeries.mkdir(finishPath);
                                this->DirsCopy(_dirPath, finishPath);
                                row.InsertDirPathToTestersMap(testerFromPasport, finishPath);
                            }
                        }
                    excelHandler.InsertRow(row);
                }
            } else qDebug() << _dirPath << " documents count == 0 =////////////////////////////////";
        } else qDebug() << _dirPath << " docx not found////////////////////////////////";
    }
}

const QString MoveHandler::FindProgramPasportPath(const QString &_dirPath) // НАХОДИТ doc-ФАЙЛ В ДИРРЕКТОРИИ
{
    const QDir dir(_dirPath);
    if (!dir.exists())
        qFatal("No such directory : %s", dir.dirName());
    else
    {
        QStringList slEntryDir = dir.entryList(QDir::Dirs);
        slEntryDir.removeOne(".");
        slEntryDir.removeOne("..");
        for (const QString& dirName : slEntryDir)
        {
            QString dirPath = dir.absolutePath();
            QString dirEntryPath = dirPath.append("/").append(dirName);
            QString KHZ = this->FindProgramPasportPath(dirEntryPath);
            if (!KHZ.isEmpty())
                return KHZ;
        }
        QStringList slEntryFiles = dir.entryList(QDir::Files);
        for (const QString& fileName : slEntryFiles)
        {
            if (fileName.contains(QRegExp(".*\.doc\w?")))
            {
                QString curDir(dir.absolutePath());
                QString docxFilePath(curDir.append("/").append(fileName));
                return docxFilePath;
            }
        }
        QString s("");
        return s;
    }
}

void MoveHandler::DirsCopy(const QString &_dirPathFrom, const QString &_dirPathTo)
{
    const QDir dirFrom(_dirPathFrom);
    if (!dirFrom.exists())
    {
        qDebug() << "No such directory : %s", dirFrom.dirName();
        return;
    }
    const QDir dirTo(_dirPathTo);
    if (!dirTo.exists())
    {
        qDebug() << "No such directory : %s", dirTo.dirName();
        return;
    }
    QStringList slEntryDirFrom = dirFrom.entryList(QDir::Dirs);
    slEntryDirFrom.removeOne("."); slEntryDirFrom.removeOne("..");
    for (const QString& dirName : slEntryDirFrom)
    {
        QString dirPathFrom = dirFrom.absolutePath();
        QString dirEntryPath = dirPathFrom.append("/").append(dirName);
        QString dirPathTo(_dirPathTo + "/" + dirName);
        QDir dir(dirEntryPath);
        dir.mkdir(dirPathTo);
        this->DirsCopy(dirEntryPath, dirPathTo);
    }
    QStringList slEntryFiles = dirFrom.entryList(QDir::Files);
    for (const QString& fileName : slEntryFiles)
    {
        QFile file(dirFrom.absolutePath() + "/" + fileName);
        file.copy(_dirPathTo + "/" + fileName);
    }
}

void MoveHandler::setPath2Kdocs(const QString &_path2Kdocs)
{
    path2Kdocs_ = _path2Kdocs;
}

void MoveHandler::setPath2Kprgs(const QString &_path2Kprgs)
{
    path2Kprgs_ = _path2Kprgs;
}

void MoveHandler::setPath2Kfrom(const QString &_path2Kfrom)
{
    path2Kfrom_ = _path2Kfrom;
}
