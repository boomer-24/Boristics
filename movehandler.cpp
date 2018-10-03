#include "movehandler.h"

MoveHandler::MoveHandler(QObject *parent) : QObject(parent)
{    
    emit this->signalStart();
    this->isExcelBusy_ = false;
}

MoveHandler::~MoveHandler()
{
    emit this->signalFinish();
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

void MoveHandler::TraverseArchiveDir()
{
    if (!this->path2Kfrom_.isEmpty())
    {
        const QDir dir(this->path2Kfrom_);
        if (!dir.exists())
            qDebug() << "No such directory : %s", dir.dirName();
        else
        {
            emit this->signalProgressToUI(0);
            QStringList slEntryDir = dir.entryList(QDir::Dirs);
            slEntryDir.removeOne(".");
            slEntryDir.removeOne("..");
            this->isExcelBusy_ = true;
            for (int i = 0; i < slEntryDir.size(); i++)
            {
                const QString& dirName = slEntryDir.at(i);
                QString dirPath = dir.absolutePath();
                QString dirEntryPath = dirPath.append("/").append(dirName);
                if (this->TraverseMainProgramDir(dirEntryPath))
                    dir.rmdir(dirEntryPath);
                emit this->signalProgressToUI(((i + 1) * 100) / slEntryDir.size());
            }
            this->isExcelBusy_ = false;
            emit this->signalInfoToUItrueTextBox("!!! ОБРАБОТКА ОФОРМИТЬ И В АРХИВ ЗАВЕРШЕНА !!!");
            emit this->signalTraverseArchiveComplete();
        }
    } else emit this->signalToUIfailTextBox("Укажи в ini.xml путь к \"ОФОРМИТЬ И В АРХИВ\"");
}

void MoveHandler::TraverseArchiveDir(const QString &_dirPath) //    ПЕРЕДАВАТЬ ПАПКУ ОФОРМИТЬ И В АРХИВ/2К
{
    if (!this->path2Kfrom_.isEmpty())
    {
        const QDir dir(_dirPath);
        if (!dir.exists())
            qDebug() << "No such directory : %s", dir.dirName();
        else
        {
            emit this->signalProgressToUI(0);
            QStringList slEntryDir = dir.entryList(QDir::Dirs);
            slEntryDir.removeOne(".");
            slEntryDir.removeOne("..");
            this->isExcelBusy_ = true;
            for (int i = 0; i < slEntryDir.size(); i++)
            {
                const QString& dirName = slEntryDir.at(i);
                QString dirPath = dir.absolutePath();
                QString dirEntryPath = dirPath.append("/").append(dirName);
                if (this->TraverseMainProgramDir(dirEntryPath))
                    dir.rmdir(dirEntryPath);
                emit this->signalProgressToUI(((i + 1) * 100) / slEntryDir.size());
            }
            this->isExcelBusy_ = false;
            emit this->signalInfoToUItrueTextBox("!!! ОБРАБОТКА ОФОРМИТЬ И В АРХИВ ЗАВЕРШЕНА !!!");
            emit this->signalTraverseArchiveComplete();
        }
    } else emit this->signalToUIfailTextBox("Укажи в ini.xml путь к \"ОФОРМИТЬ И В АРХИВ\"");
}

bool MoveHandler::TraverseMainProgramDir(const QString &_dirPath) // ДЛЯ БАЗОВОЙ ПАПКИ ОДНОЙ ПРОГРАММЫ
{
    bool markerForReturn(false);
    const QDir dir(_dirPath);
    if (!dir.exists())
        qDebug() << "No such directory : %s", dir.dirName();
    else
    {
        QString docxPath(this->FindProgramPassportPath(_dirPath));
        if (!docxPath.isEmpty())
        {
            QStringList slPath(docxPath.split("."));
            QFile docxFile(docxPath);
            QString newPath(QCoreApplication::applicationDirPath().append("/Passport").append(".").append(slPath.last()));
            if (QFile::exists(newPath))
                QFile::remove(newPath);
            docxFile.copy(newPath);
            auto ini = CoInitializeEx(NULL, COINIT_MULTITHREADED );
            //            Passport2KAnalyzer Passport(this, newPath);
            if (ini == 0 || 1)
            {
                Passport2KAnalyzer Passport;
                QObject::connect(&Passport, SIGNAL(signalInfoToUIfailTextBox(QString)),
                                 this, SIGNAL(signalToUIfailTextBox(QString)));
                if (Passport.Run(newPath))
                {
                    if (Passport.DocumentsCount())
                    {
                        Passport.Initialize();
                        ExcelHandler excelHandler(this, this->path2Kexcel_);
                        QObject::connect(&excelHandler, SIGNAL(signalSeriesNotExist(QString)),
                                         this, SIGNAL(signalToUIfailTextBox(QString)));
                        const QVector<RowInExcelTable> &vRows = Passport.rowsInExcelTable();
                        Passport.QuitWord();
                        this->PassportCopy(docxPath, vRows.first().series());
                        if (QFile::exists(newPath))
                            QFile::remove(newPath);
                        QString dirNameForMoving(_dirPath.split("/").last());
                        if (!dirNameForMoving.contains(QRegExp("\\d+\\s\\w")))
                        {
                            QRegExp rgxp("(\\d+)(\\w+)");
                            //                    int pos = rgxp.indexIn(dirNameForMoving);
                            dirNameForMoving.insert(rgxp.pos(2), " ");
                        }
                        for (RowInExcelTable row : vRows)
                        {
                            const QString &series(row.series());
                            const QStringList &slTestersFromPassport(row.slTesters());
                            const QStringList &slTestersExistingDirs(QDir(this->path2Kprgs_).entryList(QDir::Dirs));//nodotanddotdot

                            for (const QString &testerExistingDirs : slTestersExistingDirs)
                                for (const QString &testerFromPassport : slTestersFromPassport)
                                {
                                    if (testerExistingDirs.contains(testerFromPassport))
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
                                        row.InsertDirPathToTestersMap(testerFromPassport, finishPath);
                                    }
                                }
                            excelHandler.InsertRow(row);
                            QString report(row.name());
                            report.append(" Успешно перемещена и записана в таблицу");
                            markerForReturn = true;
                            emit this->signalInfoToUItrueTextBox(report);
                        }
                    } else emit this->signalToUIfailTextBox(QString(_dirPath).append(" Проблема с word documents count == 0"));
                } else emit this->signalToUIfailTextBox(QString(newPath).append(" Не создался Word.Application. Зови на помощь//MH"));
            } else emit this->signalToUIfailTextBox(QString(" Проблема с CoInitializeEx() Зови на помощь"));
            CoUninitialize();
        } else emit this->signalToUIfailTextBox(QString("В ").append(_dirPath).append(" не найден паспорт. Без паспорта не работаю"));
    }
    return markerForReturn;
}

const QString MoveHandler::FindProgramPassportPath(const QString &_dirPath) // НАХОДИТ doc-ФАЙЛ В ДИРРЕКТОРИИ
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
            QString KHZ = this->FindProgramPassportPath(dirEntryPath);
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
        emit this->signalToUIfailTextBox(QString("Нет директории ").append(dirFrom.dirName()));
        qDebug() << "No such directory : %s", dirFrom.dirName();
        return;
    }
    const QDir dirTo(_dirPathTo);
    if (!dirTo.exists())
    {
        emit this->signalToUIfailTextBox(QString("Нет директории ").append(dirTo.dirName()));
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

void MoveHandler::PassportCopy(const QString &_pathFrom, const QString &_series)
{
    QString pathTo(this->path2Kdocs_);
    QString element(_pathFrom.split(QRegExp("(\\\\)|(/)")).last());
    if (!element.contains(QRegExp("\\d+\\s\\w")))
    {
        QRegExp rgxp("(\\d+)(\\w+)");
        int pos = rgxp.indexIn(element);
        element.insert(rgxp.pos(2), " ");
    }
    pathTo.append("/").append(_series).append("/").append(element);
    QDir dirPassport(this->path2Kdocs_);
    if (dirPassport.entryList(QDir::Dirs).contains(_series))
    {
        while (QFile::exists(pathTo))
        {
            QRegExp re("\.doc\w*");
            re.indexIn(pathTo);
            int pos = re.pos(0);
            pathTo.insert(pos, "_Newer");
        }
        QFile file(_pathFrom);
        file.copy(pathTo);
    } else
    {
        dirPassport.mkdir(_series);
        QFile file(_pathFrom);
        file.copy(pathTo);
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

void MoveHandler::slotStartOperations()
{

    Initialize(QCoreApplication::applicationDirPath().append("/ini.xml"));
    TraverseArchiveDir();
}

QString MoveHandler::path2Kexcel() const
{
    return path2Kexcel_;
}

bool MoveHandler::isExcelBusy() const
{
    return isExcelBusy_;
}
