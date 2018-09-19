#ifndef EXCELHANDLER_H
#define EXCELHANDLER_H

#include "rowinexceltable.h"

#include <QAxObject>
#include <QDebug>
#include <QMap>
#include <QDate>

class ExcelHandler
{
private:
    QAxObject* excelApp_;
    QAxObject* workbooks_;
    QAxObject* workbook_;

public:
    ExcelHandler();
    ExcelHandler(const QString& _excelPath);
    ~ExcelHandler();
    void InsertRow(const RowInExcelTable& _row);
    void InsertTextToCell(QAxObject *_sheet, int _row, int _column, const QString& _text);
    void InsertHyperLinkTextToCell(QAxObject *_sheet, int _row, int _column, const QString& _address, const QString& _text);
    int FindEmptyRow(QAxObject *_sheet);
    void InsertEmptyColumn(QAxObject *_sheet, int _row, int _column, QString _text);    
    bool isFitDateCell(QAxObject *_sheet, int _row, int _column);
    QMap<QString, int> FillTestersMap(const QStringList& _slTesters, QAxObject *_sheet, int _row, int _column);
    int GetMaxValueFromMap(QList<int> _list);
    QPair<bool, bool> isFitDateFind(const QString &_dateFromTable, const QDate &_dateSelected); //first bool - isFit date, second - isFail
    QStringList getNewProgram(const QDate _dateSelected);
};

#endif // EXCELHANDLER_H
