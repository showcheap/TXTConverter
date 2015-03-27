#include "mainwindow.h"
#include "ui_mainwindow.h"
#include <QFileDialog>
#include <QDir>
#include <QFile>
#include <QDebug>
#include <QStringList>
#include <QSqlDatabase>
#include <QSqlQuery>
#include <QSqlError>
#include <QMessageBox>
#include <QSettings>

MainWindow::MainWindow(QWidget *parent) :
    QMainWindow(parent),
    ui(new Ui::MainWindow)
{
    ui->setupUi(this);
    this->programInit();
}

MainWindow::~MainWindow()
{
    delete ui;
}

// Tombol pilih folder
void MainWindow::on_btnFolder_clicked()
{
    QString folderOutput = QFileDialog::getExistingDirectory(this,"Pilih Folder Output",QDir::homePath(),QFileDialog::ShowDirsOnly);
    ui->lineFolder->setText(folderOutput);
    QSettings().setValue("lastfolder",folderOutput);

}

// Tombol Pilih file
void MainWindow::on_btnFile_clicked()
{
    QString fileInput = QFileDialog::QFileDialog::getOpenFileName(this,
                                                                  tr("Pilih File"),
                                                                  QDir::homePath(),
                                                                  tr("Text Files(*.txt *.TXT)"));
    ui->lineFile->setText(fileInput);
}

// TOmbol proses
void MainWindow::on_btnProses_clicked()
{
    QString filePath = ui->lineFile->text();
    QString folderOut = ui->lineFolder->text()+"/";

    if(filePath.isEmpty() || ui->lineFolder->text().isEmpty()){
        QMessageBox::warning(this,"Error, Tidak Boleh Kosong","Chek File source & Folder tujuan");
        return;
    }

    // Get File Source
    QFile file(filePath);
    QFileInfo fInfo(file.fileName());
    QString fileName(fInfo.fileName());
    fileName.remove(QRegExp("\.[A-z]{3,4}$"));

    QString fileOut = folderOut+fileName+".xls";


    // Buat folder dulu jika belum ada
    if(!QDir(folderOut).exists()){
        if(QDir().mkpath(folderOut)){
            qDebug()<<"INFO : Buat Folder Berhasil";
        }else{
            qDebug()<<"ERROR : Buat Folder Gagal";
        }
    }else{
        //  Check if file exist
        QFile currentFile(fileOut);
        if(currentFile.exists()){
            qDebug()<<"INFO : File sudah ada, hapus..!!";
            currentFile.setPermissions(QFile::WriteUser);
            currentFile.remove();
        }else{
            qDebug()<<"INFO: File belum ada, Membuat file baru";
        }
    }

    qDebug()<<"INFO : File Output "+fileOut;


    //  Copy xls to local

    if(QFile::copy(":/FILE/EXAMPLE.xls",fileOut)){
        //  Set read & write
        QFile tempFile(fileOut);
        tempFile.setPermissions(QFile::WriteUser);
    }else{
        qDebug()<<"ERROR : Gagal copy file xls";
    }



    if (file.open(QIODevice::ReadOnly | QIODevice::Text)){

        QSqlDatabase xlsDB = QSqlDatabase::addDatabase("QODBC");
        //  Set db ke file xls
        xlsDB.setDatabaseName("DRIVER={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};DBQ="+QString(fileOut)+";ReadOnly=0;");


        if(xlsDB.open()){

            qDebug()<<"INFO: XLS OPEN";

            QTextStream in(&file);

            int numLine = 0;

            QString field = "CREATE TABLE TRANSAKSIxx (";


            QSqlQuery q;
            q.prepare("INSERT INTO TRANSAKSIx VALUES(?,?,?,?,?,?,?,?,?,?)");

            // Bikin variabel value
            QVariantList vno,vdate,vtime,vbillid,vraceid,vamount,vbillerid,vterminalid,vrefnum,vperiod,vsettledate;

            // Buat Tabel Dulu

            while (!in.atEnd()) {
                QString line = in.readLine();
                QStringList row = line.split("|");

                // Buat tabel berdasarkan field
                if(numLine == 0){
                    // Iki mesti judul
                    QStringList fieldTransform;
                    //fieldTransform.append("DATA_NO NUMBER");
                    for(int x=0;x<=row.count();x++){
                        QString fieldName = row.value(x).trimmed();

                        if(!fieldName.isEmpty() || !fieldName.isNull()){
                            fieldTransform.append("D_"+fieldName.toUpper()+" TEXT");
                        }

                    }
                    field += fieldTransform.join(", ");
                    field += ")";

                    /*
                    QSqlQuery query;

                    if(query.exec(field)){
                        qDebug()<<"INFO : Query OK";
                    }else{
                        qDebug()<<"ERROR : Create Table error "<<query.lastError().text();
                    }
                    */
                }else{


                    vno << numLine;
                    //q.addBindValue(vno);
                    vdate << row.value(0);
                    //q.addBindValue(vdate);
                    vtime << row.value(1);
                    //q.addBindValue(vtime);
                    vbillid << row.value(2);
                    //q.addBindValue(vbillid);
                    vraceid << row.value(3);
                    //q.addBindValue(vraceid);
                    vamount << row.value(4);
                    //q.addBindValue(vamount);
                    vbillerid << row.value(5);
                    //q.addBindValue(vbillerid);
                    vterminalid << row.value(6);
                    //q.addBindValue(vterminalid);
                    vrefnum << row.value(7);
                    //q.addBindValue(vrefnum);
                    vperiod << row.value(8);
                    //q.addBindValue(vperiod);
                    vsettledate << row.value(9);
                    //q.addBindValue(vsettledate);
                    //vsettledate << row.value(10);
                    //q.addBindValue(vdate);
                }


                numLine++;

            } // endwhile
            //vno << numLine;
            //q.addBindValue(vno);
            //vdate << row.value(0);
            q.addBindValue(vdate);
            //vtime << row.value(1);
            q.addBindValue(vtime);
            //vbillid << row.value(2);
            q.addBindValue(vbillid);
            //vraceid << row.value(3);
            q.addBindValue(vraceid);
            //vamount << row.value(4);
            q.addBindValue(vamount);
            //vbillerid << row.value(5);
            q.addBindValue(vbillerid);
            //vterminalid << row.value(6);
            q.addBindValue(vterminalid);
            //vrefnum << row.value(7);
            q.addBindValue(vrefnum);
            //vperiod << row.value(8);
            q.addBindValue(vperiod);
            //vsettledate << row.value(9);
            q.addBindValue(vsettledate);
            //vdate << row.value(10);


            if(!q.execBatch()){
                qDebug()<<"ERROR : Insert error "+q.lastError().text()+"\n"+q.lastQuery();
            }else{
                qDebug()<<"INFO : Convert data sejumlah "+QString::number(numLine)+" baris.";
                QMessageBox::information(this,"Suksess...!!","Data berhasil di konversi ke file Excel\n\nLokasi File Konversi : "+fileOut);
            }
            q.clear();
        }else{
            qDebug()<<"ERROR: XLS Gak bisa dibuka"<<xlsDB.lastError();
        }

        /*
        QSqlQuery qDel;
        if(!qDel.exec("DROP TABLE IF EXISTS [TRANSAKSI$]")){
            qDebug()<<"ERROR : Hapus tabel temp "+qDel.lastError().text();
        }
        */

        xlsDB.close();

        statusBar()->showMessage("Lokasi File "+fileOut);
    }else{
        qDebug()<<"Gaiso bukak file";
    }

    QSqlDatabase::removeDatabase("qt_sql_default_connection");

}

void MainWindow::programInit(){
    QCoreApplication::setOrganizationName("MantabMedia");
    QCoreApplication::setOrganizationDomain("mantabmedia.com");
    QCoreApplication::setApplicationName("PPOBCONVERTER");

    this->setWindowTitle("PPOB Text Converter");
    ui->lineFolder->setText(QSettings().value("lastfolder",QDir::homePath()).toString());
}

void MainWindow::on_actionKeluar_triggered()
{
    this->close();
}

void MainWindow::on_actionTentang_triggered()
{
    QMessageBox::information(this,
                             "Tentang Aplikasi",
                             "<b>Aplikasi PPOB TXT to XLS Converter<\/b><br\/><br\/>Dibuat Oleh <a href='http:\/\/www.sucipto.net'>Sucipto<\/a><br\/>Ngawi, 26 Maret 2015");
}
