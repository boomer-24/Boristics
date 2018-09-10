#ifndef EXCELHANDLER_H
#define EXCELHANDLER_H

#include "rowinexceltable.h"

#include <QAxObject>
#include <QDebug>
#include <QMap>

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
    int FindEmptyRow(QAxObject *_sheet);
    void InsertEmptyColumn(QAxObject *_sheet, int _row, int _column, QString _text);    
    bool isFitDateCell(QAxObject *_sheet, int _row, int _column);
    QMap<QString, int> FillTestersMap(const QStringList& _slTesters, QAxObject *_sheet, int _row, int _column);
    int GetMaxValueFromMap(QList<int> _list);
};

#endif // EXCELHANDLER_H
