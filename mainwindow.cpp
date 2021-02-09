#include "mainwindow.h"
#include "ui_mainwindow.h"
#include <QScreen>
#include <QFileDialog>
#include <QMessageBox>

#include "xlsxdocument.h"
#include "xlsxchartsheet.h"
#include "xlsxcellrange.h"
#include "xlsxchart.h"
#include "xlsxrichstring.h"
#include "xlsxworkbook.h"
using namespace QXlsx;

MainWindow::MainWindow(QWidget *parent) : QMainWindow(parent), ui(new Ui::MainWindow) {
    ui->setupUi(this);
    this->move(QGuiApplication::primaryScreen()->geometry().center() - this->rect().center());
}

MainWindow::~MainWindow() {
    delete ui;
}

bool checkFileExists(QString filePath) {
    if (!QFile::exists(filePath)) {
        QMessageBox msgBox;
        msgBox.setWindowTitle("Error");
        msgBox.setText("File "" + filePath + "" was not found.");
        msgBox.setDefaultButton(QMessageBox::Ok);
        msgBox.setIcon(QMessageBox::Critical);
        msgBox.exec();
        return false;
    }
    return true;
}

const QStringList XlsxColumns = { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z" };

void MainWindow::on_convertBtn_clicked() {
    if (!checkFileExists(ui->filePathInput->text())) return;

    QFile file(ui->filePathInput->text());
    if (!file.open(QIODevice::ReadOnly | QIODevice::Text)) return;

    QTextStream textReader(&file);
    textReader.setEncoding(QStringConverter::Utf8);

    QXlsx::Document xlsx;

    QString header = textReader.readLine();
    QStringList headerData = header.split(",");
    for (int i = 0; i < headerData.size(); ++i) {
        if(i > XlsxColumns.length()) continue;
        xlsx.write(XlsxColumns.at(i) + "1", headerData.at(i));
    }
    // QString::number(i + 1)

    int row = 1;
    while(!textReader.atEnd()) {
        row++;
        QString line = textReader.readLine();
        QStringList lineData = line.split(",");
        for (int i = 0; i < lineData.size(); ++i) {
            if(i > XlsxColumns.length()) continue;
            xlsx.write(XlsxColumns.at(i) + QString::number(row), lineData.at(i));
        }
    }

    QString fileName = QFileDialog::getSaveFileName(this, tr("Save File"), ui->filePathInput->text() + ".xlsx", tr("Xlsx (*.xlsx)"));
    if (!fileName.isNull() && !fileName.isEmpty()) {
         xlsx.saveAs(fileName);
    }
}

void MainWindow::on_convertTableBtn_clicked() {
    if (!checkFileExists(ui->filePathInput->text())) return;

}

void MainWindow::on_selectFileBtn_clicked() {
    auto fileName = QFileDialog::getOpenFileName(this, tr("Open CSV"), "", tr("Nubank CSV (*.csv)"));
    ui->filePathInput->setText(fileName);
}
