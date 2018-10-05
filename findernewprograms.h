#ifndef FINDERNEWPROGRAMS_H
#define FINDERNEWPROGRAMS_H

#include "excelhandler.h"
#include <QObject>
#include <combaseapi.h>

class FinderNewPrograms : public QObject
{
    Q_OBJECT
public:
    explicit FinderNewPrograms(QObject *parent = 0);

    bool isExcelBusy() const;

signals:
    void signalInfoToUInewProgramTrueTextBox(QString info);
    void signalInfoToUInewProgramFailTextBox(QString infoErr);
    void signalSearchNewProgramComplete();
    void signalCurrentSheetToUI(QString);

public slots:
    void slotStartOperation(const QString &_excelPath, const QDate &_date);

private:
    bool isExcelBusy_;
};

#endif // FINDERNEWPROGRAMS_H
