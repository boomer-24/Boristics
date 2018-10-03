#ifndef ROWINEXCELTABLE_H
#define ROWINEXCELTABLE_H

#include <QStringList>
#include <QMap>

class RowInExcelTable
{
private:
    QString series_, name_, KY_, author_, TY_, TYcorrection_, date_;
    QStringList slTesters_;
    QMap<QString, QString> mapTesterAndDirPath_;
public:
    RowInExcelTable();
    RowInExcelTable(const QString& _series, const QString& _name, const QString& _KY,
                    const QString& _author, const QString& _TY, const QString& _TYcorrection, const QString& _date);
    RowInExcelTable(const QString& _series, const QString& _name, const QString& _KY,
                    const QString& _author, const QString& _TY, const QString& _TYcorrection,
                    const QString& _date, const QStringList& _slTesters);
    void InsertDirPathToTestersMap(const QString& _tester, const QString& _dirPath);
    QMap<QString, QString> mapTesterAndDirPath() const;
    QString series() const;
    QString name() const;
    QString KY() const;
    QString author() const;
    QString TY() const;
    QString TYcorrection() const;
    QString date() const;
    const QStringList &slTesters() const;
};

#endif // ROWINEXCELTABLE_H
