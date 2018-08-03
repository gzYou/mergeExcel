#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QMainWindow>
#include "qt_windows.h"
#include <QtWidgets>
#include <QFileDialog>

namespace Ui {
class MainWindow;
}

class MainWindow : public QMainWindow
{
    Q_OBJECT

public:
    explicit MainWindow(QWidget *parent = 0);
    ~MainWindow();
    void getALLExcelProperties();
    QVariant castListListVariant2Variant(const QList<QList<QVariant> > &cells);

private slots:
    void merge();
    void addTo();

    void display(int state);
private:
    Ui::MainWindow *ui;
    QList<QList<QString>> properties; //保存各Excel表字段.properties[0]-体质人类学表型特征录入文件;properties[1]-骨密度;properties[2]-体成分;properties[3]-检验数据
    QList<QVariant> chinese;//英文-中文对照
    QList<QVariant> partName;//分区名称
    QString filePath[5];//存储文件位置
    QMap<QVariant,int> keyword; //用来存储各表统一编号(关键字)
    int start[4] = {0,165,186,282};
    int end[4]={164,185,281,309};
    int keywordPos[4]={1,0,0,1};//关键字在各表字段中的位置
    int firstData[4]={2,1,1,2};//第一条有效数据所在行号
    int namePos[4]={2,1,1,0};//姓名所在位置
    QString savePath = "C:\\excel\\汇总.xlsx"; //汇总表存储位置
    QMap<QString,int> fileIndex;
    QList<QString> file;
};

#endif // MAINWINDOW_H
