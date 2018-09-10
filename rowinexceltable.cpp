#include "rowinexceltable.h"

RowInExcelTable::RowInExcelTable()
{

}

RowInExcelTable::RowInExcelTable(const QString &_series, const QString &_name, const QString &_KY,
                                 const QString &_author, const QString &_TY, const QString &_TYcorrection,
                                 const QString &_date)
{
    this->series_ = _series;
    this->name_ = _name;
    this->KY_ = _KY;
    this->author_ = _author;
    this->TY_ = _TY;
    this->TYcorrection_ = _TYcorrection;
    this->date_ = _date;
}

RowInExcelTable::RowInExcelTable(const QString &_series, const QString &_name, const QString &_KY,
                                 const QString &_author, const QString &_TY, const QString &_TYcorrection,
                                 const QString &_date, const QStringList& _slTesters)
{
    this->series_ = _series;
    this->name_ = _name;
    this->KY_ = _KY;
    this->author_ = _author;
    this->TY_ = _TY;
    this->TYcorrection_ = _TYcorrection;
    this->date_ = _date;
    this->slTesters_ = _slTesters;
}

QString RowInExcelTable::series() const
{
    return series_;
}

QString RowInExcelTable::name() const
{
    return name_;
}

QString RowInExcelTable::KY() const
{
    return KY_;
}

QString RowInExcelTable::author() const
{
    return author_;
}

QString RowInExcelTable::TY() const
{
    return TY_;
}

QString RowInExcelTable::TYcorrection() const
{
    return TYcorrection_;
}

QString RowInExcelTable::date() const
{
    return date_;
}

const QStringList &RowInExcelTable::slTesters() const
{
    return slTesters_;
}
