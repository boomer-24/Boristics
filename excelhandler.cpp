#include "excelhandler.h"

ExcelHandler::ExcelHandler()
{

}

ExcelHandler::ExcelHandler(const QString &_excelPath)
{
    this->excelApp_ = new QAxObject("Excel.Application", 0);
    this->workbooks_ = excelApp_->querySubObject("Workbooks");
    this->workbook_ = workbooks_->querySubObject("Open(const QString&)", _excelPath);
}

ExcelHandler::~ExcelHandler()
{
    this->workbook_->dynamicCall("Save()");
    this->workbook_->dynamicCall("Close()");
    this->workbooks_->dynamicCall("Close()");
    this->excelApp_->dynamicCall("Quit()");
    delete this->excelApp_;
}

void ExcelHandler::InsertRow(const RowInExcelTable &_row)
{        
    QAxObject* sheets = this->workbook_->querySubObject("Worksheets");
    int countSheets = sheets->property("Count").toInt();
    int columnName = 0, columnCondition = 0, columnKY = 0, columnDeveloper = 0, columnTY = 0, columnTYcorrection = 0, columnDate = 0;

    QString series = _row.series();
    bool foundSeries(false);
    for (int i = 1; i <= countSheets; i++)
    {
        QAxObject* sheet = sheets->querySubObject("Item(int)", i);
        if (sheet->property("Name").toString() == series)
        {
            foundSeries = true;
            QAxObject* usedRange = sheet->querySubObject("UsedRange");
            QAxObject* columns = usedRange->querySubObject("Columns");
            int countColumns = columns->property("Count").toInt();
            QMap<QString, int> mapTestersColumns;
            for (int rowCounter = 1; rowCounter < 4; rowCounter++)
                for (int columnCounter = 1; columnCounter < countColumns; columnCounter++)
                {
                    QAxObject* cell = sheet->querySubObject("Cells(int,int)", rowCounter, columnCounter);
                    QVariant value = cell->property("Value");
                    QString strValue(value.toString().toUtf8());
                    if (strValue.contains(QRegExp(".*азван")))
                        columnName = columnCounter;
                    else if (strValue.contains(QRegExp(".*остояние")))
                        columnCondition = columnCounter;
                    else if (strValue.contains("КУ") || strValue.contains("KY"))
                        columnKY = columnCounter;
                    else if (strValue.contains(QRegExp(".*втор")))
                        columnDeveloper = columnCounter;
                    else if (strValue.contains(QRegExp(".*аписано")))
                        columnTY = columnCounter;
                    else if (strValue.contains(QRegExp(".*оррекц.*ТУ")))
                        columnTYcorrection = columnCounter;
                    else if (strValue.contains("Дата") && isFitDateCell(sheet, rowCounter, columnCounter))
                        columnDate = columnCounter;
                    else if (strValue.contains(QRegExp(".*аспростр.*")))
                        mapTestersColumns = this->FillTestersMap(_row.slTesters(), sheet, rowCounter + 1, columnCounter);
                }
            int rowForInsertion = this->FindEmptyRow(sheet);
            if (rowForInsertion)
            {
                this->InsertTextToCell(sheet, rowForInsertion, columnName, _row.name());
                this->InsertTextToCell(sheet, rowForInsertion, columnCondition, "Используется");
                this->InsertTextToCell(sheet, rowForInsertion, columnKY, _row.KY());
                this->InsertTextToCell(sheet, rowForInsertion, columnDeveloper, _row.author());
                this->InsertTextToCell(sheet, rowForInsertion, columnTY, _row.TY());
                this->InsertTextToCell(sheet, rowForInsertion, columnTYcorrection, _row.TYcorrection());
                this->InsertTextToCell(sheet, rowForInsertion, columnDate, _row.date());
                for (QString key : mapTestersColumns.keys())
                {
                    if (_row.slTesters().contains(key))
                        this->InsertHyperLinkTextToCell(sheet,
                                                        rowForInsertion, mapTestersColumns.value(key),
                                                        _row.mapTesterAndDirPath().value(key), "+");
//                        this->InsertTextToCell(sheet, rowForInsertion, mapTestersColumns.value(key), "+");
                    else
                        this->InsertTextToCell(sheet, rowForInsertion, mapTestersColumns.value(key), "-");
                }
            }
        }
    } if (!foundSeries)
        qDebug() << "series " << series << " not exist";
}

void ExcelHandler::InsertTextToCell(QAxObject* _sheet, int _row, int _column, const QString &_text)
{
    if (_column)
    {
        QAxObject* cellIns = _sheet->querySubObject("Cells(int,int)", _row, _column);
        cellIns->setProperty("Value", QVariant(_text));
    }
}

void ExcelHandler::InsertHyperLinkTextToCell(QAxObject *_sheet, int _row, int _column, const QString &_address, const QString &_text)
{
    QString alphabet(" ABCDEFGHIJKLMNOPQRSTUVWXYZ");
    QString request("Range(\"");
    request.append(alphabet.at(_column)).append(QString::number(_row)).append("\")");

    QAxObject* hLinks = _sheet->querySubObject("HyperLinks");
    QAxObject* sheetRange = _sheet->querySubObject(request.toStdString().data());
    QAxObject* link = hLinks->querySubObject("Add(Object, String))", sheetRange->asVariant(), _address);
    link->setProperty("TextToDisplay", "+");

}

int ExcelHandler::FindEmptyRow(QAxObject *_sheet)
{
    QAxObject* usedRange = _sheet->querySubObject("UsedRange");
    QAxObject* columns = usedRange->querySubObject("Columns");
    int countColumns = columns->property("Count").toInt();
    QAxObject* rows = usedRange->querySubObject("Rows");
    int countRows = rows->property("Count").toInt();
    int emptyRowNumber = 0;

    for (int row = 1; row <= countRows; row++)
    {
        bool rowIsEmpty(true);
        QAxObject* cellDate = _sheet->querySubObject("Cells(int,int)", row, 1);
        QVariant valueDate = cellDate->property("Value");
        if (valueDate.toString().isEmpty())
        {
            for (int column = 1; column <= countColumns; column++)
            {
                QAxObject* cellDate = _sheet->querySubObject("Cells(int,int)", row, column);
                QVariant valueDate = cellDate->property("Value");
                if (!valueDate.toString().isEmpty())
                {
                    rowIsEmpty = false;
                    break;
                }
            }
            if (rowIsEmpty)
            {
                emptyRowNumber = row;
                break;
            }
        } else continue;
    }
    if (!emptyRowNumber)
    {
        emptyRowNumber = rows->property("Count").toInt() + 1;
    }
    return emptyRowNumber;
}

void ExcelHandler::InsertEmptyColumn(QAxObject *_sheet, int _row, int _column, QString _text)
{
    QString alphabet(" ABCDEFGHIJKLMNOPQRSTUVWXYZ");
    QString request("Range(\"");
    request.append(alphabet.at(_column)).append(QString::number(_row)).append("\")");
    QAxObject* sheetRange = _sheet->querySubObject(request.toStdString().data());

    //    QAxObject* sheetRange = _sheet->querySubObject("Range(\"C9\")");
    QAxObject* entireColumn = sheetRange->querySubObject("EntireColumn");
    entireColumn->dynamicCall("Insert");
    QAxObject* cellIns = _sheet->querySubObject("Cells(int,int)", _row, _column);
    cellIns->setProperty("Value", QVariant(_text));
}

bool ExcelHandler::isFitDateCell(QAxObject *_sheet, int _row, int _column)
{
    if (_row > 1)
        _row--;
    for (int r = _row; r >= 1; r--)
        for (int c = _column; c > _column - 4; c--)
        {
            QAxObject* cell = _sheet->querySubObject("Cells(int,int)", r, c);
            QVariant value = cell->property("Value");
            QString s(value.toString().toUtf8());
            if ((s.contains(QRegExp(".*оррекц.*"))))
                return false;
        }
    return true;
}

QMap<QString, int> ExcelHandler::FillTestersMap(const QStringList &_slTesters, QAxObject *_sheet, int _row, int _column)
{
    QMap<QString, int> mapTestersColumns;
    QAxObject* usedRange = _sheet->querySubObject("UsedRange");
    QAxObject* columns = usedRange->querySubObject("Columns");
    int countColumns = columns->property("Count").toInt();
    QRegExp rx("\\b(\\d+)\\b");

    for (int c = _column; c <= countColumns; c++)
    {
        QAxObject* cell = _sheet->querySubObject("Cells(int,int)", _row, c);
        QVariant value = cell->property("Value");
        QString tester(value.toString().toUtf8());
        int pos = rx.indexIn(tester);
        if (pos != -1)
            mapTestersColumns.insert(tester, c);
        else break;
    }
    for (const QString& tester : _slTesters)
    {
        if (mapTestersColumns.contains(tester))
            continue;
        else
        {
            int col = this->GetMaxValueFromMap(mapTestersColumns.values());
            col++;// т.к. вставка столбца происходит слева
            this->InsertEmptyColumn(_sheet, _row, col, tester);
            mapTestersColumns.insert(tester, col);
        }
    }
    return mapTestersColumns;
}

int ExcelHandler::GetMaxValueFromMap(QList<int> _list)
{
    int max(0);
    for (int value : _list)
    {
        if (value > max)
            max = value;
    }
    return max;
}

QPair<bool, bool> ExcelHandler::isFitDateFind(const QString &_dateFromTable, const QDate &_dateSelected)
{
    if (!_dateFromTable.isEmpty())
    {
        QStringList slDayMonthYear(_dateFromTable.split("-"));
        if (slDayMonthYear.size() == 3)
        {
            QString year = slDayMonthYear.at(0);
            QString month = slDayMonthYear.at(1);
            QString day = slDayMonthYear.at(2);
            if (day.size() > 2)
                day.truncate(2);
            QDate dateFromTable(year.toInt(), month.toInt(), day.toInt());
            if (dateFromTable >= _dateSelected)
                return QPair<bool, bool> (true, false);
            else return QPair<bool, bool> (false, false);

        } else return QPair<bool, bool> (false, true);
    } else return QPair<bool, bool> (false, false);
}

QStringList ExcelHandler::getNewProgram(const QDate _dateSelected)
{
    if (!excelApp_->isNull())
    {
        QStringList slNamesElements;
        QAxObject* sheets = this->workbook_->querySubObject("Worksheets");
        int countSheets = sheets->property("Count").toInt();
        for (int i = 1; i < countSheets; i++)
        {
            int column = 1;
            QAxObject* sheet = sheets->querySubObject("Item(int)", i);
            QAxObject* usedRange = sheet->querySubObject("UsedRange");
            QAxObject* columns = usedRange->querySubObject("Columns");
            int countColumns = columns->property("Count").toInt();
            QAxObject* cell = sheet->querySubObject("Cells(int,int)", 2, column);
            QVariant value = cell->property("Value");

            bool next(false);
            while (!value.toString().toUtf8().contains("Дата"))
            {
                column++;
                if (column <= countColumns)
                {
                    cell = sheet->querySubObject("Cells(int,int)", 2, column);
                    value = cell->property("Value");
                } else
                {
                    next = true;
                    break;
                }
            }
            if (next) continue;
            QAxObject* rows = usedRange->querySubObject("Rows");
            int countRows = rows->property("Count").toInt();
            for (int row = 3; row < countRows; row++)
            {
                QAxObject* cellDate = sheet->querySubObject("Cells(int, int)", row, column);
                QVariant valueDate = cellDate->property("Value");
                if (this->isFitDateFind(valueDate.toString(), _dateSelected).first)
                {
                    QAxObject* cellName = sheet->querySubObject("Cells(int,int)", row, 1);
                    QVariant valueName = cellName->property("Value");
                    slNamesElements.push_back(valueName.toString());
                    qDebug() << valueName.toString();
                } else if (this->isFitDateFind(valueDate.toString(), _dateSelected).second)
                {
                    QAxObject* cellName = sheet->querySubObject("Cells(int,int)", row, 1);
                    QVariant valueName = cellName->property("Value");
                    qDebug() << "fail with " << valueName << "!!!!!!!!!!!!!!!!!!";
                }
            }
        }
        return slNamesElements;
    }
}
