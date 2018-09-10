#ifndef PASPORT2KANALYZER_H
#define PASPORT2KANALYZER_H

#include "rowinexceltable.h"

#include <QAxObject>
#include <QDebug>
#include <QRegExp>
#include <QDate>

class Pasport2KAnalyzer
{
private:
    QAxObject *WordApp_;
    QAxObject *Documents_;
    QAxObject *Table_;
    QString  docxPath_;

    QString firstPage_;

    bool developerMarker_ = false;    

    QStringList slTesters_, slNames_;
    QString TY_, TYcorrection_, KY_, developer_, series_, name_;
    QString content_;
    QStringList slContent_;
    QStringList slContentText_;

public:
    Pasport2KAnalyzer();
    Pasport2KAnalyzer(const QString &_docxPath);
    ~Pasport2KAnalyzer();
    void QuitWord();
    void SetDocxPath(const QString& _docxPath);
    void Initialize();
    void EjectOsnastka();
    void EjectDeveloperFromTable();
    void EjectDeveloperFromPlainText();
    void EjectTesters();
    void EjectSeries();    
    void EjectTYandCorrection();
    const QStringList &testers() const;
    const QString &TY() const;
    const QString &TYcorrection() const;
    const QString &KY() const;
    const QString &developer() const;
    const QVector<RowInExcelTable> rowsInExcelTable() const;
    int DocumentsCount() const;
    void PrepareNameInsertSpace(QString& _name);
};

#endif // PASPORT2KANALYZER_H
