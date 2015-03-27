#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QMainWindow>

namespace Ui {
class MainWindow;
}

class MainWindow : public QMainWindow
{
    Q_OBJECT

public:
    explicit MainWindow(QWidget *parent = 0);
    ~MainWindow();

private slots:
    void on_btnFolder_clicked();

    void on_btnFile_clicked();

    void on_btnProses_clicked();

    void on_actionKeluar_triggered();

    void on_actionTentang_triggered();

private:
    Ui::MainWindow *ui;

    void programInit();
};

#endif // MAINWINDOW_H
