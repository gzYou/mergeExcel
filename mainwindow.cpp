#include "mainwindow.h"
#include "ui_mainwindow.h"
#include "excelengine.h"
#include "qdebug.h"
#include "qt_windows.h"
#include <QString>
#include <QStringList>
#include <QFileDialog>
#include <QTableWidget>
#include<algorithm>
#include<QList>
using namespace std;


MainWindow::MainWindow(QWidget *parent) :
    QMainWindow(parent),
    ui(new Ui::MainWindow)
{
    ui->setupUi(this);
    getALLExcelProperties();
    QObject::connect(ui->pushButton,SIGNAL(clicked()),this,SLOT(merge()));
    QObject::connect(ui->addTo,SIGNAL(clicked()),this,SLOT(addTo()));


    QObject::connect(ui->checkBox,SIGNAL(stateChanged(int)),this,SLOT(display(int)));
    QObject::connect(ui->checkBox_2,SIGNAL(stateChanged(int)),this,SLOT(display(int)));
    QObject::connect(ui->checkBox_3,SIGNAL(stateChanged(int)),this,SLOT(display(int)));
    QObject::connect(ui->checkBox_4,SIGNAL(stateChanged(int)),this,SLOT(display(int)));
    QObject::connect(ui->checkBox_5,SIGNAL(stateChanged(int)),this,SLOT(display(int)));
}

MainWindow::~MainWindow()
{
    delete ui;
}


/*自定义排序规则
 * bool cmp()
*/
bool cmp(QList<QVariant> &l1,QList<QVariant> &l2)
{
    return l1[1]<l2[1];
}

/*获取各Excel字段名称
 * getAllExcelProperties()
 * properties[0]-体质人类学表型特征录入文件;
 * properties[1]-骨密度;
 * properties[2]-体成分;
 * properties[3]-检验数据;
*/
void MainWindow::getALLExcelProperties()
{
    QList<QString> temp; //用来暂存各Excel属性
    temp.clear();
    temp<<"序号"<<"编号"<<"姓名"<<"性别"<<"民族"<<"居住时间"<<"身份证号"<<"出生日期"<<"调查日期"<<"联系电话"<<"教育年限"<<"1题"<<"2题"<<"3题"<<"4题"<<"5题"<<"6题"<<"7题"<<"8题"<<"9题"<<"10题"<<"11题"<<"12题"<<"13题"<<"14题"<<"15题①"<<"15题②"<<"15题③"<<"15题④"<<"16题①"<<"16题②"<<"16题③"<<"16题④"<<"17题"<<"18题"<<"19题"<<"20题"<<"21题"<<"22题"<<"发旋"<<"发形"<<"额发际"<<"额倾斜"<<"眉毛"<<"眉弓"<<"上眼睑"<<"蒙古褶"<<"眼裂高"<<"眼倾斜"<<"鼻根高"<<"鼻背观"<<"鼻基部"<<"鼻孔径"<<"颧突度"<<"耳尖"<<"耳垂"<<"唇侧面"<<"美人沟"<<"下颏"<<"尖舌"<<"卷舌"<<"翻舌"<<"叠舌"<<"利手"<<"扣手"<<"交叉臂"<<"拇指"<<"头长"<<"头宽"<<"两耳外宽"<<"乳突间宽"<<"额宽"<<"耳屏间宽"<<"面宽"<<"下颌角宽"<<"容貌面高"<<"眼内宽"<<"眼外宽"<<"形态面高"<<"鼻高"<<"鼻长"<<"鼻翼高"<<"鼻宽"<<"鼻深"<<"鼻下颏高"<<"唇皮高"<<"唇高"<<"红唇厚"<<"口宽"<<"耳长"<<"耳宽"<<"肱骨径"<<"股骨径"<<"身高"<<"体重"<<"收缩压"<<"舒张压"<<"耳屏高"<<"颏下高"<<"肩峰高"<<"胸上高"<<"桡骨高"<<"茎突高"<<"指尖高"<<"髂前棘"<<"胫上高"<<"指距"<<"坐高"<<"手宽"<<"足长"<<"足宽"<<"胸宽"<<"肩宽"<<"骨盆宽"<<"内踝下"<<"内踝地面高"<<"座椅地面高"<<"头围"<<"颈围"<<"胸围"<<"吸气围"<<"呼气围"<<"腰围"<<"腹围"<<"臀围"<<"大腿围"<<"小腿围"<<"上臂围"<<"臂缩围"<<"前臂围"<<"面颊褶"<<"三头褶"<<"二头褶"<<"肩胛褶"<<"髂上褶"<<"髂前褶"<<"小腿褶"<<"是否釆血"<<"是否照相"<<"凳高"<<"心电图心率"<<"心电图诊断1"<<"心电图诊断2"<<"心电图诊断3"<<"心电图诊断4"<<"心电图诊断5"<<"B超诊断1"<<"B超诊断2"<<"B超诊断3"<<"B超诊断4"<<"B超诊断5"<<"握力（左）"<<"握力（右）"<<"血糖（空）"<<"血糖（餐)"<<"备注"<<"父亲民族"<<"母亲民族"<<"环食"<<"三叶舌"<<"指甲"<<"门齿"<<"鼻梁"<<"利足"<<"鼻翼宽";
    properties.push_back(temp);
    temp.clear();
    temp<<"卡号"<<"姓名"<<"性别"<<"出生日期"<<"年龄"<<"身份证号"<<"民族"<<"地区"<<"籍贯"<<"城乡工作种类"<<"家庭住址"<<"注册日期"<<"测试时间"<<"测试部位"<<"T值"<<"T值_评价"<<"%年轻成人"<<"Z值"<<"%同龄人"<<"骨强度指数"<<"骨折风险分析";
    properties.push_back(temp);
    temp.clear();
    temp<<"卡号"<<"姓名"<<"性别"<<"出生日期"<<"年龄"<<"身份证号"<<"民族"<<"地区"<<"籍贯"<<"城乡工作种类"<<"家庭住址"<<"注册日期"<<"测试时间"<<"类型"<<"身高"<<"体重"<<"身体质量指数"<<"身体质量指数_评价"<<"去脂体重"<<"脂肪量"<<"肌肉量"<<"推定骨量"<<"身体水分"<<"蛋白质"<<"细胞内液"<<"细胞外液"<<"体脂肪率"<<"体脂肪率_评价"<<"肌肉量比值"<<"肌肉量比值_评价"<<"内脏脂肪等级"<<"内脏脂肪等级_评价"<<"内脏脂肪面积"<<"内脏脂肪含量"<<"皮下脂肪含量"<<"腰臀比"<<"腰臀比_评价"<<"基础代谢"<<"总能量代谢"<<"基础代谢分析_评价"<<"浮肿指数"<<"浮肿指数_评价"<<"躯干肌肉量"<<"躯干肌肉量等级_评价"<<"左上肢肌肉量"<<"左上肢肌肉量等级_评价"<<"左下肢肌肉量"<<"左下肢肌肉量等级_评价"<<"右上肢肌肉量"<<"右上肢肌肉量等级_评价"<<"右下肢肌肉量"<<"右下肢肌肉量等级_评价"<<"躯干脂肪量"<<"躯干脂肪率"<<"躯干脂肪率等级_评价"<<"左上肢脂肪量"<<"左上肢脂肪率"<<"左上肢脂肪率等级_评价"<<"左下肢脂肪量"<<"左下肢脂肪率"<<"左下肢脂肪率等级_评价"<<"右上肢脂肪量"<<"右上肢脂肪率"<<"右上肢脂肪率等级_评价"<<"右下肢脂肪量"<<"右下肢脂肪率"<<"右下肢脂肪率等级_评价"<<"目标体重"<<"体重控制"<<"脂肪控制"<<"肌肉控制"<<"综合评分"<<"右臂R_5kHz"<<"右臂R_50kHz"<<"右臂R_250kHz"<<"右臂R_500kHz"<<"左臂R_5kHz"<<"左臂R_50kHz"<<"左臂R_250kHz"<<"左臂R_500kHz"<<"身体左侧R_5kHz"<<"身体左侧R_50kHz"<<"身体左侧R_250kHz"<<"身体左侧R_500kHz"<<"右腿R_5kHz"<<"右腿R_50kHz"<<"右腿R_250kHz"<<"右腿R_500kHz"<<"左腿R_5kHz"<<"左腿R_50kHz"<<"左腿R_250kHz"<<"左腿R_500kHz"<<"双腿R_5KHz"<<"双腿R_50KHz"<<"双腿R_250KHz"<<"双腿R_500KHz";
    properties.push_back(temp);
    temp.clear();
    temp<<"姓名"<<"病历号"<<"病人类型"<<"性别"<<"年"<<"龄"<<"送检日期"<<"操作员"<<"报告日期"<<"标本"<<"BUN"<<"CREA"<<"UA"<<"ALT"<<"AST"<<"AST/ALT"<<"GGT"<<"TBIL"<<"DBIL"<<"IDBIL"<<"TP"<<"ALB"<<"GLO"<<"A/G"<<"CHOL"<<"TG"<<"HDL"<<"LDL";
    properties.push_back(temp);
    chinese<<"尿素氮"<<"肌酐"<<"尿酸"<<"谷丙转氨酶"<<"谷草转氨酶"<<"谷丙谷草比值/"<<"谷氨酸氨基转移酶"<<"总胆红素"<<"直接胆红素"<<"间接胆红素 "<<"总蛋白"<<"白蛋白"<<"球蛋白"<<"白球比值/"<<"总胆固醇"<<"甘油三酯"<<"高密度脂蛋白"<<"低密度脂蛋白";
    partName<<"基本情况"<<"第一部分        问卷部分          第一部分        问卷部分"<<"第二部分  传统形态观察类表型         第二部分  传统形态观察类表型      第二部分  传统形态观察类表型"<<"第三部分1          头面部测量                      第三部分1          头面部测量"<<"第三部分2   肢体测量     第三部分2  肢体测量"<<"第三部分3 体围测量     第三部3分 体围测量"<<"遗传学观察目标";
}

/*将QList<QList<QVariant>>数据类型的数据转换为QVariant类型的数据
 * castListListVariant2Variant(const QList<QList<QVariant> > &cells)
*/
QVariant MainWindow::castListListVariant2Variant(const QList<QList<QVariant> > &cells)
{
    QVariantList vars;
    const int rows = cells.size();
    for(int i=0;i<rows;++i)
    {
        vars.append(QVariant(cells[i]));
    }
    return QVariant(vars);
}

/*文件合并槽函数
 * merge()
*/
void MainWindow::merge()
{
    QFile f(savePath);
    if(f.exists())
    {
        QMessageBox* box = new QMessageBox;
        box->setWindowTitle("Notice");
        box->setText("汇总文件已存在，请点击追加");
        box->show();
        return;
    }
    ui->pushButton->setEnabled(false);
    QApplication::processEvents();


    keyword.insert(QVariant("test"),-1);
    qDebug()<<savePath;
    QAxObject* pExcel = new QAxObject("Excel.Application");//连接Excel控件
    pExcel->dynamicCall("SetVisible (bool Visible)",false);//不显示窗体
    pExcel->setProperty("DisplayAlerts",false);//不显示任何警告信息
    QAxObject* pWorkbooks = pExcel->querySubObject("WorkBooks");//获取工作簿
    pWorkbooks->dynamicCall("Add"); //添加一个新的工作簿
    QAxObject* pWorkbook = pExcel->querySubObject("ActiveWorkBook");//新建一个xlsx
    QAxObject* pWorksheet = pWorkbook->querySubObject("WorkSheets(int)",1);//打开第一个sheet

    /*合并单元格*/
    QAxObject* merge_range = pWorksheet->querySubObject("Range(const QString)","A1:k1");
    QAxObject* interior = merge_range->querySubObject("Interior");
    interior->setProperty("Color",QColor(127,255,212));
    merge_range->setProperty("MergeCells",true);
    merge_range->dynamicCall("SetValue(const QVariant&)",QVariant(partName[0]));


    merge_range = pWorksheet->querySubObject("Range(const QString)","L1:AM1");
    interior = merge_range->querySubObject("Interior");
    interior->setProperty("Color",QColor(186,85,211));
    merge_range->setProperty("MergeCells",true);
    merge_range->dynamicCall("SetValue(const QVariant&)",QVariant(partName[1]));

    merge_range = pWorksheet->querySubObject("Range(const QString)","AN1:BO1");
    interior = merge_range->querySubObject("Interior");
    interior->setProperty("Color",QColor(192,192,192));
    merge_range->setProperty("MergeCells",true);
    merge_range->dynamicCall("SetValue(const QVariant&)",QVariant(partName[2]));

    merge_range = pWorksheet->querySubObject("Range(const QString)","BP1:CO1");
    interior = merge_range->querySubObject("Interior");
    interior->setProperty("Color",QColor(255,255,0));
    merge_range->setProperty("MergeCells",true);
    merge_range->dynamicCall("SetValue(const QVariant&)",QVariant(partName[3]));

    merge_range = pWorksheet->querySubObject("Range(const QString)","CP1:DM1");
    interior = merge_range->querySubObject("Interior");
    interior->setProperty("Color",QColor(0,255,255));
    merge_range->setProperty("MergeCells",true);
    merge_range->dynamicCall("SetValue(const QVariant&)",QVariant(partName[4]));

    merge_range = pWorksheet->querySubObject("Range(const QString)","DN1:EG1");
    interior = merge_range->querySubObject("Interior");
    interior->setProperty("Color",QColor(192,192,192));
    merge_range->setProperty("MergeCells",true);
    merge_range->dynamicCall("SetValue(const QVariant&)",QVariant(partName[5]));

    merge_range = pWorksheet->querySubObject("Range(const QString)","FA1:FI1");
    interior = merge_range->querySubObject("Interior");
    interior->setProperty("Color",QColor(192,192,192));
    merge_range->setProperty("MergeCells",true);
    merge_range->dynamicCall("SetValue(const QVariant&)",QVariant(partName[6]));

    QList<QVariant> headList;//汇总表表头

    /*将四个Excel的字段导入表头并建立表头*/
    for(int i=0;i<4;i++)
    {
        for(int j=0;j<properties[i].length();j++)
            headList<<properties[i][j];
    }
    QAxObject* range = pWorksheet->querySubObject("Range(const QString)","KG1:KX1");
    range->dynamicCall("SetValue(const QVariant&)",QVariant(chinese));
    range = pWorksheet->querySubObject("Range(const QString)","A2:KX2");
    range->dynamicCall("SetValue(const QVariant&)",QVariant(headList));

    /*将四个表的数据导入*/
    QList<QList<QVariant>> res;
    ExcelEngine* excel[4];
    QString address="C:\\excel\\";

    QList<QList<QVariant>> transport;
    QList<QVariant> tempLine;
    int rowCount = 1;

    for(int i=0;i<4;i++)
    {
        excel[i]=new ExcelEngine(address+filePath[i]);
    }
    for(int i=0;i<4;i++)
    {
        excel[i]->Open();
        excel[i]->castVariant2ListListVariant();
        res.clear();
        res = excel[i]->getRes();
        for(int j=firstData[i];j<res.length();j++)
        {
            if(keyword[res[j][keywordPos[i]]]) //如果该keyword已经存在，在所在行对应分区补入数据
            {
                int pos = keyword[res[j][keywordPos[i]]]-1;
//                qDebug()<<"存在"<<res[j][keywordPos[i]]<<",位于"<<keyword[res[j][keywordPos[i]]];
                //range = pWorksheet->querySubObject("Range(const QString )", start[i]+QString::number(pos)+":"+end[i]+QString::number(pos));//定位到插入位置
                for(int k=0;k<res[j].length();k++)
                {
                    transport[pos][k+start[i]] = res[j][k];
                }


            }
            else if(QVariant(res[j][keywordPos[i]]).toString()!="") //如果数据不存在且数据有效（存在关键字），加入transport
            {
//                qDebug()<<"不存在"<<res[j][keywordPos[i]];
                tempLine.clear();
                for(int k=0;k<start[i];k++)
                {
                    tempLine<<"";
                }
                for(int k=0;k<res[j].length();k++)
                {
                    tempLine<<res[j][k];
                }
                for(int k=end[i]+1;k<310;k++)
                {
                    tempLine<<"";
                }
                tempLine[1]=tempLine[start[i]+keywordPos[i]];
                transport.push_back(tempLine);
//                qDebug()<<"已加入"<<res[j][keywordPos[i]]<<"位置:"<<rowCount;
                keyword.insert(res[j][keywordPos[i]],rowCount++);
            }
        }
        excel[i]->Close();
        delete excel[i];
    }
    sort(transport.begin(),transport.end(),cmp);
    /*将所有在transpo暂存的数据转换格式存入汇总表相应分区*/
    int newRows = transport.length();
    range = pWorksheet->querySubObject("Range(const QString )","A3:KX"+QString::number(2+newRows));
    range->setProperty("Value", castListListVariant2Variant(transport));
//    qDebug()<<"合并时的res长度:"<<2+newRows;

    /*格式*/
    range = pWorksheet->querySubObject("Range(const QString )","G3:G"+QString::number(2+newRows));
    range->dynamicCall("SetNumberFormatLocal", "##################");
    range = pWorksheet->querySubObject("Range(const QString )","FO3:FO"+QString::number(2+newRows));
    range->dynamicCall("SetNumberFormatLocal", "##################");
    range = pWorksheet->querySubObject("Range(const QString )","GJ3:GJ"+QString::number(2+newRows));
    range->dynamicCall("SetNumberFormatLocal", "##################");

    range = pWorksheet->querySubObject("Range(const QString )","H3:I"+QString::number(2+newRows));
    range->setProperty("NumberFormatLocal", "yyyy/m/d");
    range = pWorksheet->querySubObject("Range(const QString )","FM3:FM"+QString::number(2+newRows));
    range->setProperty("NumberFormatLocal", "yyyy/m/d");
    range = pWorksheet->querySubObject("Range(const QString )","GH3:GH"+QString::number(2+newRows));
    range->setProperty("NumberFormatLocal", "yyyy/m/d");
    range = pWorksheet->querySubObject("Range(const QString )","KC3:KC"+QString::number(2+newRows));
    range->setProperty("NumberFormatLocal", "yyyy/m/d");
    range = pWorksheet->querySubObject("Range(const QString )","KE3:KE"+QString::number(2+newRows));
    range->setProperty("NumberFormatLocal", "yyyy/m/d");
    range = pWorksheet->querySubObject("Range(const QString )","FU3:FV"+QString::number(2+newRows));
    range->setProperty("NumberFormatLocal", "yyyy/m/d hh:mm");
    range = pWorksheet->querySubObject("Range(const QString )","GP3:GQ"+QString::number(2+newRows));
    range->setProperty("NumberFormatLocal", "yyyy/m/d hh:mm");



    pWorkbook->dynamicCall("SaveAs(const QString&)",QDir::toNativeSeparators(savePath));//保存至filepath，注意一定要用QDir::toNativeSeparators将路径中的"/"转换为"\"，不然一定保存不了。
    pWorkbook->dynamicCall("Close()");//关闭工作簿
    pExcel->dynamicCall("Quit()");//关闭excel

    QMessageBox* box = new QMessageBox;
    box->setWindowTitle("Notice");
    box->setText("汇总完成!");
    box->show();

    ui->pushButton->setEnabled(true);


}

/*数据追加
 * addTo()
*/

void MainWindow::addTo()
{
    /*检查原始汇总文件是否存在*/
    QFile f(savePath);
    if(!f.exists())
    {
        QMessageBox* box = new QMessageBox;
        box->setWindowTitle("Notice");
        box->setText("原始汇总文件不存在,请点击汇总.");
        box->show();
        return;
    }

    ui->addTo->setEnabled(false);
    QApplication::processEvents();

    /*打开汇总表，读入已有编号信息并保存*/
    ExcelEngine mergedExcel("C:\\excel\\汇总.xlsx");
    mergedExcel.Open();
    mergedExcel.castVariant2ListListVariant();
    QList<QList<QVariant>> Res = mergedExcel.getRes();
//    qDebug()<<"已有res的长度:"<<Res.length();
    keyword.clear();
    for(int i=2;i<Res.length();i++)
    {
//        qDebug()<<QVariant(Res[i][1]).toString();
        keyword.insert(Res[i][1],i);
    }
    mergedExcel.Close();
    qDebug()<<"打开成功"<<endl;

    QString address="C:\\excel\\";
    ExcelEngine* excel[4];
    for(int i=0;i<4;i++)
    {
        excel[i] = new ExcelEngine(address+filePath[i]);
    }

    /*按顺序处理四个源表
     * 首先，判断是否是追加的数据
     * 其次，如果是要追加的数据，判断该追加数据是否被其他源表已经加到汇总表中
    */
    QMap<QVariant,int> keywordAdded;//存储追加数据的统一编号(关键字)
    QList<QList<QVariant>> transport,res;
    QList<QVariant> tempLine;
    int rowCount = 1;

    for(int i=0;i<4;i++)
    {
        excel[i]->Open();
        excel[i]->castVariant2ListListVariant();
        res.clear();
        res = excel[i]->getRes();
        for(int j=firstData[i];j<res.length();j++)
        {
            if(!keyword[res[j][keywordPos[i]]]) //新ID人员
            {
                if(keywordAdded[res[j][keywordPos[i]]]) //如果该keyword已经存在，在所在行对应分区补入数据
                {
                    int pos = keywordAdded[res[j][keywordPos[i]]]-1;
                    for(int k=0;k<res[j].length();k++)
                    {
                        transport[pos][k+start[i]] = res[j][k];
                    }


                }
                else if(QVariant(res[j][keywordPos[i]]).toString()!="") //如果数据不存在且数据有效（存在关键字），加入transport
                {
                    tempLine.clear();
                    for(int k=0;k<start[i];k++)
                    {
                        tempLine<<"";
                    }
                    for(int k=0;k<res[j].length();k++)
                    {
                        tempLine<<res[j][k];
                    }
                    for(int k=end[i]+1;k<310;k++)
                    {
                        tempLine<<"";
                    }
                    tempLine[1]=tempLine[start[i]+keywordPos[i]];
                    transport.push_back(tempLine);
                    keywordAdded.insert(res[j][keywordPos[i]],rowCount++);
                }
            }
            else //如果该ID存在，那么检查对应分区是否有数据，检查依据：姓名
            {
                int pos = keyword[res[j][keywordPos[i]]];
                if(QVariant(Res[pos][start[i]+namePos[i]]).toString()=="")//该源文件在汇总文件中的对应分区空白，则补录
                {
                    for(int k=0;k<res[j].length();k++)
                    {
                        Res[pos][start[i]+k] = res[j][k];
                    }
                }

            }
        }
        excel[i]->Close();
        delete excel[i];
    }
    /*排序时要考虑到所有数据*/
    for(int i=0;i<transport.length();i++)
    {
        Res.push_back(transport[i]);
    }
    sort(Res.begin()+2,Res.end(),cmp);
    int newRows = Res.length();
    QAxObject* pExcel = new QAxObject("Excel.Application");
    QAxObject* pWorkbooks = pExcel->querySubObject("WorkBooks"); //获取工作簿
    QAxObject* pWorkbook = pWorkbooks->querySubObject("Open(QString, QVariant)","C:\\excel\\汇总.xlsx",QVariant(0)); //打开xls对应的工作簿
    QAxObject* pWorksheet = pWorkbook->querySubObject("WorkSheets(int)", 1);//打开第一个sheet
    QAxObject* range = pWorksheet->querySubObject("Range(const QString )","A1:KX"+QString::number(newRows));
    range->setProperty("Value", castListListVariant2Variant(Res));

    /*格式*/
    range = pWorksheet->querySubObject("Range(const QString )","G3:G"+QString::number(newRows));
    range->dynamicCall("SetNumberFormatLocal", "##################");
    range = pWorksheet->querySubObject("Range(const QString )","FO3:FO"+QString::number(newRows));
    range->dynamicCall("SetNumberFormatLocal", "##################");
    range = pWorksheet->querySubObject("Range(const QString )","GJ3:GJ"+QString::number(newRows));
    range->dynamicCall("SetNumberFormatLocal", "##################");

    range = pWorksheet->querySubObject("Range(const QString )","H3:I"+QString::number(newRows));
    range->setProperty("NumberFormatLocal", "yyyy/m/d");
    range = pWorksheet->querySubObject("Range(const QString )","FM3:FM"+QString::number(newRows));
    range->setProperty("NumberFormatLocal", "yyyy/m/d");
    range = pWorksheet->querySubObject("Range(const QString )","GH3:GH"+QString::number(newRows));
    range->setProperty("NumberFormatLocal", "yyyy/m/d");
    range = pWorksheet->querySubObject("Range(const QString )","KC3:KC"+QString::number(newRows));
    range->setProperty("NumberFormatLocal", "yyyy/m/d");
    range = pWorksheet->querySubObject("Range(const QString )","KE3:KE"+QString::number(newRows));
    range->setProperty("NumberFormatLocal", "yyyy/m/d");
    range = pWorksheet->querySubObject("Range(const QString )","FU3:FV"+QString::number(newRows));
    range->setProperty("NumberFormatLocal", "yyyy/m/d hh:mm");
    range = pWorksheet->querySubObject("Range(const QString )","GP3:GQ"+QString::number(newRows));
    range->setProperty("NumberFormatLocal", "yyyy/m/d hh:mm");


    pWorkbook->dynamicCall("Save()");
    pWorkbook->dynamicCall("Close(bool)", true);
    pExcel->dynamicCall("Quit()");

    QMessageBox* box = new QMessageBox;
    box->setWindowTitle("Notice");
    box->setText("追加完成!");
    box->show();
    ui->addTo->setEnabled(true);
}

/*数据浏览
 * display()
*/
/*void MainWindow::display(int state)
{
    QCheckBox *act=qobject_cast<QCheckBox*>(sender());//使用Qt的类型转换，将指针恢复为QAction类型
    QString fileName = ((QCheckBox*)act)->text();
    QString address =  "C:\\excel\\"+filePath[tabIndex[fileName]];
    QFile f(address);
    if(state == Qt::Checked)
    {
        if(!f.exists())
        {
            QMessageBox* box = new QMessageBox;
            box->setWindowTitle("Notice");
            box->setText("文件不存在！");
            box->show();
            return;
        }
        ui->tabWidget->setTabText(tabIndex[fileName],act->text());
        ExcelEngine excel(address);
        excel.Open();
        excel.ReadDataToTable(tableW[fileName]);
        excel.Close();
    }
    if(state == Qt::Unchecked)
    {
        tableW[fileName]->clear();
        tableW[fileName]->setColumnCount(0);
        tableW[fileName]->setRowCount(0);
        ui->tabWidget->setTabText(tabIndex[fileName],"");
    }
}
*/

void MainWindow::display(int state)
{
    QCheckBox *act=qobject_cast<QCheckBox*>(sender());//使用Qt的类型转换，将指针恢复为QAction类型
    QString fileName = ((QCheckBox*)act)->text();
    qDebug()<<fileName;
    QString address =  "C:\\excel\\"+fileName+".xlsx";
    QFile f(address);

    if(state == Qt::Checked)
    {
        if(!f.exists())
        {
            QMessageBox* box = new QMessageBox;
            box->setWindowTitle("Notice");
            box->setText("文件不存在！");
            box->show();
            act->setCheckState(Qt::Unchecked);
            return;
        }

        QTableWidget *tableWidget = new QTableWidget();
        ExcelEngine excel(address);
        excel.Open();
        excel.ReadDataToTable(tableWidget);
        excel.Close();
        tableWidget->setEditTriggers(QAbstractItemView::NoEditTriggers);


        ui->tabWidget->insertTab(0,tableWidget,fileName);
        file.insert(0,fileName);
        fileIndex.insert(fileName,0);
        for(int i = 1;i<file.length();i++)
        {
            fileIndex[file[i]] += 1;
        }
        ui->tabWidget->setCurrentIndex(0);

    }
    else
    {
        if(fileIndex.find(fileName) == fileIndex.end()) return;
        ui->tabWidget->removeTab(fileIndex[fileName]);
        int index = file.indexOf(fileName);
        for(int i = index + 1;i<file.length();i++)
        {
            fileIndex[file[i]] -=1;
        }
        file.removeAt(index);
        fileIndex.remove(fileName);

    }
}

