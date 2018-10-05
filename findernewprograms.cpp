#include "findernewprograms.h"

FinderNewPrograms::FinderNewPrograms(QObject *parent) : QObject(parent)
{
    this->isExcelBusy_ = false;
}

void FinderNewPrograms::slotStartOperation(const QString& _excelPath, const QDate& _date)
{
    auto ini = CoInitializeEx(NULL, COINIT_MULTITHREADED );
    if (ini == 0 || 1)
    {
        this->isExcelBusy_ = true;
        ExcelHandler EH(this, _excelPath);
        QObject::connect(&EH, SIGNAL(signalInfoToUInewProgramTextBox(QString)),
                         this, SIGNAL(signalInfoToUInewProgramTrueTextBox(QString)));
        QObject::connect(&EH, SIGNAL(signalInfoToUInewProgramFailTextBox(QString)),
                         this, SIGNAL(signalInfoToUInewProgramFailTextBox(QString)));
        QObject::connect(&EH, SIGNAL(signalCurrentSheetToUI(QString)),
                         this, SIGNAL(signalCurrentSheetToUI(QString)));
        EH.getNewProgram(_date);
        emit this->signalSearchNewProgramComplete();
        this->isExcelBusy_ = false;
    } else emit this->signalInfoToUInewProgramFailTextBox(QString(" Проблема с CoInitializeEx() Зови на помощь"));
}

bool FinderNewPrograms::isExcelBusy() const
{
    return isExcelBusy_;
}
