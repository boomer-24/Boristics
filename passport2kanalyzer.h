#ifndef Passport2KANALYZER_H
#define Passport2KANALYZER_H

#include "rowinexceltable.h"

#include <QObject>
#include <QAxObject>
#include <QDebug>
#include <QRegExp>
#include <QDate>
#include <combaseapi.h>

class Passport2KAnalyzer : public QObject
{
    Q_OBJECT
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
    Passport2KAnalyzer(QObject *parent = 0);
    Passport2KAnalyzer(QObject *parent, const QString &_docxPath);
    ~Passport2KAnalyzer();
    bool Run(const QString &_docxPath);
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

signals:
    void signalInfoToUIfailTextBox(const QString& _info);
};

#endif // Passport2KANALYZER_H
