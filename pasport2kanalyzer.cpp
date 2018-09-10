#include "pasport2kanalyzer.h"

Pasport2KAnalyzer::Pasport2KAnalyzer()
{

}

Pasport2KAnalyzer::Pasport2KAnalyzer(const QString &_docxPath)
{
    this->WordApp_ = new QAxObject("Word.Application");
    if (!this->WordApp_)
        qDebug() << "WordApp fall";
    this->Documents_ = WordApp_->querySubObject("Documents()");
    if (!this->Documents_)
        qDebug() << "Documents fall";
    this->docxPath_ = _docxPath;
    this->Documents_->querySubObject("Open(const QString)", this->docxPath_);
}

Pasport2KAnalyzer::~Pasport2KAnalyzer()
{    
    ////    if (this->Table_)
    ////        delete this->Table_;
    //    //    if (this->Documents_)
    //    //        delete this->Documents_;
    // НИГДЕ НЕТ CLOSE ДЛЯ WORDа
}

void Pasport2KAnalyzer::QuitWord()
{
    this->WordApp_->dynamicCall("Quit()");
//    if (this->WordApp_)
//        delete this->WordApp_;
}

void Pasport2KAnalyzer::SetDocxPath(const QString &_docxPath)
{
    if (!_docxPath.isEmpty())
    {
        this->WordApp_ = new QAxObject("Word.Application");
        if (!this->WordApp_)
            qDebug() << "WordApp fall";
        this->Documents_ = WordApp_->querySubObject("Documents()");
        if (!this->Documents_)
            qDebug() << "Documents fall";
        this->docxPath_ = _docxPath;
        qDebug() << this->Documents_->querySubObject("Open(const QString)", this->docxPath_);

        qDebug() << this->Documents_->property("Count").toInt();
    }
}

void Pasport2KAnalyzer::Initialize()
{
    QAxObject *ActiveDocument = this->Documents_->querySubObject("Item(1)");
    QAxObject *Tables = ActiveDocument->querySubObject("Tables");
    int tablesCount = Tables->property("Count").toInt();
    QAxObject *myRange = ActiveDocument->querySubObject("Content()");
    this->content_ = myRange->property("Text").toString();
    this->firstPage_ = this->content_.split("\f").at(0);
    this->slContent_ = this->content_.split(" ");
    this->EjectSeries();
    for (int tableNumber = 1; tableNumber <= tablesCount; tableNumber++)
    {
        bool marker = false;
        QString query = QString("Tables(%1)").arg(tableNumber);
        const char *tableQuery = query.toStdString().c_str();
        this->Table_ = ActiveDocument->querySubObject(tableQuery);
        if (!this->Table_)
        {
            qDebug() << tableNumber << "   table no open";
            return;
        }
        int intCols = this->Table_->querySubObject("Range")->querySubObject("Columns")->property("Count").toInt();
        int intRows = this->Table_->querySubObject("Range")->querySubObject("Rows")->property("Count").toInt();
        for (int row = 1; row <= 2; row++)
        {
            for (int column = 1; column <= intCols; column++)
            {
                auto Cell = this->Table_->querySubObject("Cell(Row, Column)", row, column);
                if (Cell)
                {
                    auto CellRange = Cell->querySubObject("Range()");
                    QString smth = CellRange->property("Text").toString();
                    if (smth.contains("снастк"))
                    {
                        this->EjectOsnastka();
                        marker = true;
                        break;
                    }
                    if (smth.contains("нженер") || smth.contains("категор") || smth.contains("прог"))
                    {
                        this->EjectDeveloperFromTable();
                        marker = true;
                        break;
                    }
                    if (smth.contains("естер"))
                    {
                        this->EjectTesters();
                        marker = true;
                        break;
                    }
                }
            }
            if (marker) break;
        }
    }
    this->EjectTYandCorrection();
    if (!this->developerMarker_)
        this->EjectDeveloperFromPlainText();
}

void Pasport2KAnalyzer::EjectOsnastka()
{
    int intCols = this->Table_->querySubObject("Range")->querySubObject("Columns")->property("Count").toInt();
    int intRows = this->Table_->querySubObject("Range")->querySubObject("Rows")->property("Count").toInt();
    int numberColumn = 0, KYrow = 0;

    for (int row = 1; row <= intRows; row++)
        for (int column = 1; column <= intCols; column++)
        {
            auto Cell = this->Table_->querySubObject("Cell(Row, Column)", row, column);
            if (Cell)
            {
                auto CellRange = Cell->querySubObject("Range()");
                QString smth = CellRange->property("Text").toString();
                if (smth.contains("номер"))
                    numberColumn = column;
                if (smth.contains("КУ"))
                    KYrow = row;
            }
        }
    if (numberColumn != 0 && KYrow != 0)
    {
        QString tempKY(this->Table_->querySubObject("Cell(Row, Column)", KYrow, numberColumn)->
                       querySubObject("Range()")->property("Text").toString());
        int indexOfSlash = tempKY.indexOf("\r");
        if (indexOfSlash != -1)
            tempKY.remove(indexOfSlash, tempKY.size() - indexOfSlash);
        this->KY_ = tempKY;
    }
}

void Pasport2KAnalyzer::EjectDeveloperFromTable()
{
    int intRows = this->Table_->querySubObject("Range")->querySubObject("Rows")->property("Count").toInt();
    int intCols = this->Table_->querySubObject("Range")->querySubObject("Columns")->property("Count").toInt();
    if (intRows >= 1 && intCols >= 2)
    {
        auto Cell = this->Table_->querySubObject("Cell(Row, Column)", 1, 2);
        if (Cell)
        {
            auto CellRange = Cell->querySubObject("Range()");
            QString devName = CellRange->property("Text").toString();
            QRegExp re1("\\.*[А-Я][а-я]+");
            QRegExp re2("[А-Я]([а-я]+)\\s*[А-Я].[А-Я].");
            if (re1.indexIn(devName) != -1)
                this->developer_ = re1.cap(0);
            if (re2.indexIn(devName) != -1)
                this->developer_ = re2.cap(0);
            this->developerMarker_ = true;
        }
    }
}

void Pasport2KAnalyzer::EjectDeveloperFromPlainText()
{
    if (!this->content_.isEmpty())
    {
        this->slContentText_ = this->content_.split("\r");        
        QString finishRow;
        QRegExp re1("([А-Я]\\.[А-Я]\\.\\s*)([А-Я][а-я]+)");
        QRegExp re2("([А-Я][а-я]+)\\s*([А-Я]\\.[А-Я]\\.)");
        for (int i = this->slContentText_.size() - 1; i > 0; i--)
        {
            QString row = this->slContentText_.at(i);
            if ((row.contains("жен") && row.contains("прогр")) || (row.contains(re1)) && row.contains(re2))
            {
                finishRow = row;
                break;
            }
        }
        if (!finishRow.isEmpty())
        {
            if (re1.indexIn(finishRow) != -1)
                this->developer_ = re1.cap(2);
            if (re2.indexIn(finishRow) != -1)
                this->developer_ = re2.cap(1);
        }
    }
}

void Pasport2KAnalyzer::EjectTesters()
{
    int intCols = this->Table_->querySubObject("Range")->querySubObject("Columns")->property("Count").toInt();
    int intRows = this->Table_->querySubObject("Range")->querySubObject("Rows")->property("Count").toInt();

    for (int row = 1; row <= 1/*intRows*/; row++)
        for (int column = 1; column <= intCols; column++)
        {
            auto Cell = this->Table_->querySubObject("Cell(Row, Column)", row, column);
            if (Cell)
            {
                auto CellRange = Cell->querySubObject("Range()");
                QString cellValue = CellRange->property("Text").toString();
                cellValue.remove("\007");
                cellValue.remove(" ");
                QStringList slCell = cellValue.split("\r");
                slCell.removeAll("");
                slCell.removeAll(" ");
                for (int i = 0; i < slCell.size(); i++)
                {
                    cellValue = slCell.at(i);
                    if (cellValue.isEmpty())
                        continue;
                    //                    if (cellValue.at(0).isDigit() && cellValue.at(1).isDigit())
                    if (cellValue.contains(QRegExp("\\d+")))
                    {
                        auto CellBellow = this->Table_->querySubObject("Cell(Row, Column)", row + 1, column);
                        if (CellBellow)
                        {
                            auto CellRangeBellow = CellBellow->querySubObject("Range()");
                            QString smthBellow = CellRangeBellow->property("Text").toString();
                            if (smthBellow.toLower().contains("pin"))
                                this->slTesters_.push_back(cellValue);
                        }
                    }
                }
            }
        }
    this->slTesters_.removeDuplicates();
}

void Pasport2KAnalyzer::EjectSeries()
{
    if (!this->firstPage_.isEmpty())
    {
        QString preparedFirstString;
        for (int pos = this->firstPage_.indexOf(QRegExp("провер*")); pos < this->firstPage_.size(); pos++)
        {
            preparedFirstString.append(this->firstPage_.at(pos));
        }
        preparedFirstString.remove(QRegExp("проверк.\\s*"));

        if (preparedFirstString.contains(QRegExp("(\\r)")))
        {
            QStringList slElements = preparedFirstString.split(QRegExp("(\\r)"));
            slElements.removeAll("");
            for (int i = 0; i < slElements.size(); i++)
            {
                if (slElements.at(i).contains(","))
                {
                    QStringList slLast = slElements.at(i).split(",");
                    for (QString name : slLast)
                    {
                        this->PrepareNameInsertSpace(name);
                        this->slNames_.push_back(name);
                    }
                } else
                    if (!slElements.at(i).isEmpty())
                    {
                        QString name(slElements.at(i));
                        this->PrepareNameInsertSpace(name);
                        this->slNames_.push_back(name);
                    }
            }
        } else
        {
            QStringList slLast = preparedFirstString.split(",");
            for (int i = 0; i < slLast.size(); i++)
            {
                QString name(slLast.at(i));
                this->PrepareNameInsertSpace(name);
                this->slNames_.push_back(name);
            }
        }
        QString name = this->slNames_.first();
        bool check(false);
        for (int i = 0; i < name.size(); i++)
        {
            if (check)
                break;
            if (name.at(i).isDigit())
            {
                for (int j = i; j < name.size(); j++)
                    if (name.at(j).isDigit())
                        this->series_.append(name.at(j));
                    else
                    {
                        check = true;
                        break;
                    }
            }
        }
    }
}

void Pasport2KAnalyzer::EjectTYandCorrection()
{
    if (!this->slContent_.isEmpty())
    {
        QString StrForProcessing;
        bool checkStart(false);
        bool checkFinish(false);
        int start = 0, finish = 0;
        for (int i = 0; i < this->slContent_.size(); i++)
        {
            if (this->slContent_.at(i).contains("соответств") && this->slContent_.at(i + 1).contains(QRegExp("[а-я]*")))
            {
                start = i + 2;
                checkStart = true;
            }
            //            else
            //                if (this->slContent_.at(i).contains("техническ") && this->slContent_.at(i).contains("услов"))
            //                {
            //                    start = i + 1;
            //                    checkStart = true;
            //                }
            if (checkStart)
            {
                for (int j = i; j < this->slContent_.size(); j++)
                {
                    if (this->slContent_.at(j).contains("(далее"))
                    {
                        finish = j + 2;
                        checkFinish = true;
                        break;
                    }
                }
            }
            if (checkFinish) break;
        }
        for (int i = start; i < finish; i++)
        {
            StrForProcessing.append(this->slContent_.at(i)).append(" ");
        }
        QString tempTY;
        int i = 0;
        while (!tempTY.contains(QRegExp("\\(\\w*\\s*\\w*\\)")) && i < StrForProcessing.size())
        {            
            tempTY.append(StrForProcessing.at(i++));           
        }
        QRegExp regexp("(\\w*\\d*)\\.(\\d*)\\.(\\d*)");
        regexp.indexIn(tempTY);        
        tempTY = regexp.cap(0);
        this->TY_ = tempTY;
        this->TY_.remove((QRegExp("\\(\\w*\\s*\\w*\\)")));
        bool checkCorrection = false;
        for (int i = 0; i < StrForProcessing.size(); i++)
        {
            if ((StrForProcessing.at(i) == "о" && StrForProcessing.at(i + 1) == "р") ||
                    (StrForProcessing.at(i) == "з" && StrForProcessing.at(i + 1) == "м"))
            {
                checkCorrection = true;
            }
            if (StrForProcessing.at(i).isDigit() && checkCorrection)
                this->TYcorrection_.append(StrForProcessing.at(i));
        }
    }
}

const QStringList &Pasport2KAnalyzer::testers() const
{
    return this->slTesters_;
}

const QString &Pasport2KAnalyzer::TY() const
{
    return this->TY_;
}

const QString &Pasport2KAnalyzer::TYcorrection() const
{
    return this->TYcorrection_;
}

const  QString &Pasport2KAnalyzer::KY() const
{
    return this->KY_;
}

const QString &Pasport2KAnalyzer::developer() const
{
    return this->developer_;
}

const QVector<RowInExcelTable> Pasport2KAnalyzer::rowsInExcelTable() const
{
    QVector<RowInExcelTable> vectorForExport;
    for (int i = 0; i < this->slNames_.size(); i++)
    {
        RowInExcelTable row(this->series_, this->slNames_.at(i), this->KY_, this->developer_,
                            this->TY_, this->TYcorrection_, QDate::currentDate().toString("dd.MM.yyyy"), this->slTesters_);
        vectorForExport.push_back(row);
    }
    return vectorForExport;
}

int Pasport2KAnalyzer::DocumentsCount() const
{
    return this->Documents_->property("Count").toInt();
}

void Pasport2KAnalyzer::PrepareNameInsertSpace(QString &_name)
{
    if (!_name.contains(QRegExp("\\d+\\s\\w")))
    {
        QRegExp rgxp("(\\d+)(\\w+)");
        int pos = rgxp.indexIn(_name);
        _name.insert(rgxp.pos(2), " ");
    }
}
