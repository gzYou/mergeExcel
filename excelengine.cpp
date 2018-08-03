#include "excelengine.h"
#include "qt_windows.h"
#include <QMessageBox>
#include<QTimer>
#include <QtWidgets>

ExcelEngine::ExcelEngine()
{
    pExcel     = NULL;
    pWorkbooks = NULL;
    pWorkbook  = NULL;
    pWorksheet = NULL;

    sXlsFile     = "";
    nRowCount    = 0;
    nColumnCount = 0;
    nStartRow    = 0;
    nStartColumn = 0;

    bIsOpen     = false;
    bIsValid    = false;
    bIsANewFile = false;
    bIsSaveAlready = false;

    HRESULT r = OleInitialize(0);
    if (r != S_OK && r != S_FALSE)
    {
        qDebug("Qt: Could not initialize OLE (error %x)", (unsigned int)r);
    }
}

ExcelEngine::ExcelEngine(QString xlsFile)
{
    pExcel     = NULL;
    pWorkbooks = NULL;
    pWorkbook  = NULL;
    pWorksheet = NULL;

    sXlsFile     = xlsFile;
    nRowCount    = 0;
    nColumnCount = 0;
    nStartRow    = 0;
    nStartColumn = 0;

    bIsOpen     = false;
    bIsValid    = false;
    bIsANewFile = false;
    bIsSaveAlready = false;

    HRESULT r = OleInitialize(0);
    if (r != S_OK && r != S_FALSE)
    {
        qDebug("Qt: Could not initialize OLE (error %x)", (unsigned int)r);
    }
}

ExcelEngine::~ExcelEngine()
{
    if ( bIsOpen )
    {
        //����ǰ���ȱ������ݣ�Ȼ��ر�workbook
        Close();
    }
    OleUninitialize();
}


/**
  *@brief ��sXlsFileָ����excel����
  *@return true : �򿪳ɹ�
  *        false: ��ʧ��
  */
bool ExcelEngine::Open(UINT nSheet, bool visible)
{

    if ( bIsOpen )
    {
        //return bIsOpen;
        Close();
    }

    nCurrSheet = nSheet;
    bIsVisible = visible;

    if ( NULL == pExcel )
    {
        pExcel = new QAxObject("Excel.Application");
        if ( pExcel )
        {
            bIsValid = true;
        }
        else
        {
            bIsValid = false;
            bIsOpen  = false;
            return bIsOpen;
        }

        pExcel->dynamicCall("SetVisible(bool)", bIsVisible);
    }

    if ( !bIsValid )
    {
        bIsOpen  = false;
        return bIsOpen;
    }

    if ( sXlsFile.isEmpty() )
    {
        bIsOpen  = false;
        return bIsOpen;
    }

    /*���ָ����ļ������ڣ�����Ҫ�½�һ��*/
    QFile f(sXlsFile);
    if (!f.exists())
    {
        bIsANewFile = true;
    }
    else
    {
        bIsANewFile = false;
    }

    if (!bIsANewFile)
    {
        pWorkbooks = pExcel->querySubObject("WorkBooks"); //��ȡ������
        pWorkbook = pWorkbooks->querySubObject("Open(QString, QVariant)",sXlsFile,QVariant(0)); //��xls��Ӧ�Ĺ�����
    }
    else
    {
        pWorkbooks = pExcel->querySubObject("WorkBooks");     //��ȡ������
        pWorkbooks->dynamicCall("Add");                       //���һ���µĹ�����
        pWorkbook  = pExcel->querySubObject("ActiveWorkBook"); //�½�һ��xls
    }

    pWorksheet = pWorkbook->querySubObject("WorkSheets(int)", nCurrSheet);//�򿪵�һ��sheet

    //�����Ѵ򿪣���ʼ��ȡ��Ӧ����
    QAxObject *usedrange = pWorksheet->querySubObject("UsedRange");//��ȡ��sheet��ʹ�÷�Χ����
    var = usedrange->dynamicCall("Value");
    QAxObject *rows = usedrange->querySubObject("Rows");
    QAxObject *columns = usedrange->querySubObject("Columns");

    //��Ϊexcel���Դ��������������ݶ���һ���Ǵ�0,0��ʼ�����Ҫ��ȡ�������±�
    nStartRow    = usedrange->property("Row").toInt();    //��һ�е���ʼλ��
    nStartColumn = usedrange->property("Column").toInt(); //��һ�е���ʼλ��

    nRowCount    = rows->property("Count").toInt();       //��ȡ����
    nColumnCount = columns->property("Count").toInt();    //��ȡ����

    bIsOpen  = true;
    return bIsOpen;
}

/**
  *@brief Open()�����غ���
  */
bool ExcelEngine::Open(QString xlsFile, UINT nSheet, bool visible)
{
    sXlsFile = xlsFile;
    nCurrSheet = nSheet;
    bIsVisible = visible;

    return Open(nCurrSheet,bIsVisible);
}

/**
  *@brief ���������ݣ�������д���ļ�
  */
void ExcelEngine::Save()
{
    if ( pWorkbook )
    {
        if (bIsSaveAlready)
        {
            return ;
        }

        if (!bIsANewFile)
        {
            pWorkbook->dynamicCall("Save()");
        }
        else /*������ĵ����½������ģ���ʹ�����ΪCOM�ӿ�*/
        {
            pWorkbook->dynamicCall("SaveAs (const QString&,int,const QString&,const QString&,bool,bool)",
                      sXlsFile,56,QString(""),QString(""),false,false);

        }

        bIsSaveAlready = true;
    }
}

/**
  *@brief �ر�ǰ�ȱ������ݣ�Ȼ��رյ�ǰExcel COM���󣬲��ͷ��ڴ�
  */
void ExcelEngine::Close()
{
    //�ر�ǰ�ȱ�������
    Save();

    if ( pExcel && pWorkbook )
    {
        pWorkbook->dynamicCall("Close(bool)", true);
        pExcel->dynamicCall("Quit()");

        delete pExcel;
        pExcel = NULL;

        bIsOpen     = false;
        bIsValid    = false;
        bIsANewFile = false;
        bIsSaveAlready = true;
    }
}

/**
  *@brief ��tableWidget�е����ݱ��浽excel��
  *@param tableWidget : ָ��GUI�е�tablewidgetָ��
  *@return ����ɹ���� true : �ɹ�
  *                  false: ʧ��
  */
bool ExcelEngine::SaveDataFrTable(QTableWidget *tableWidget)
{
    if ( NULL == tableWidget )
    {
        return false;
    }
    if ( !bIsOpen )
    {
        return false;
    }

    int tableR = tableWidget->rowCount();
    int tableC = tableWidget->columnCount();

    //��ȡ��ͷд����һ��
    for (int i=0; i<tableC; i++)
    {
        if ( tableWidget->horizontalHeaderItem(i) != NULL )
        {
            this->SetCellData(1,i+1,tableWidget->horizontalHeaderItem(i)->text());
        }
    }

    //д����
    for (int i=0; i<tableR; i++)
    {
        for (int j=0; j<tableC; j++)
        {
            if ( tableWidget->item(i,j) != NULL )
            {
                this->SetCellData(i+2,j+1,tableWidget->item(i,j)->text());
            }
        }
    }

    //����
    Save();

    return true;
}

void ExcelEngine::castVariant2ListListVariant()
{
    res.clear();
    QVariantList varRows = var.toList();
    if(varRows.isEmpty())
        {
            return;
        }
    const int rowCount = varRows.size();
    QVariantList rowData;
    for(int i=0;i<rowCount;++i)
    {
        rowData = varRows[i].toList();
        res.push_back(rowData);
    }
}

/**
  *@brief ��ָ����xls�ļ��а����ݵ��뵽tableWidget��
  *@param tableWidget : ִ��Ҫ���뵽��tablewidgetָ��
  *@return ����ɹ���� true : �ɹ�
  *                   false: ʧ��
  */
bool ExcelEngine::ReadDataToTable(QTableWidget *tableWidget)
{
    if ( NULL == tableWidget )
    {
        qDebug()<<"tableWidget error!";
        return false;
    }
    castVariant2ListListVariant();
    tableWidget->setEditTriggers(QAbstractItemView::NoEditTriggers);

    //�Ȱ�table���������
    int tableColumn = tableWidget->columnCount();
    tableWidget->clear();
    for (int n=0; n<tableColumn; n++)
    {
        tableWidget->removeColumn(0);
    }

//    int rowcnt    = nStartRow + nRowCount;
//    int columncnt = nStartColumn + nColumnCount;

    //��ȡexcel�еĵ�һ��������Ϊ��ͷ
    QStringList headerList;
    for (int n = 0; n<res[0].length(); n++ )
    {
        headerList<<res[0][n].toString();
    }
    //���´�����ͷ
    tableWidget->setColumnCount(nColumnCount);
    tableWidget->setHorizontalHeaderLabels(headerList);


    //����������
    if(res.length()>0)
    {
        for (int i = 1, r=0; i < res.length(); i++,r++ )  //��
        {
            tableWidget->insertRow(r); //��������
            for (int j = 0; j < nColumnCount; j++ )  //��
            {
                tableWidget->setItem(r,j,new QTableWidgetItem(res[i][j].toString()));
            }
        }

    }

    return true;
}


/**
  *@brief ��ȡָ����Ԫ�������
  *@param row : ��Ԫ����к�
  *@param column : ��Ԫ����к�
  *@return [row,column]��Ԫ���Ӧ������
  */
QVariant ExcelEngine::GetCellData(UINT row, UINT column)
{
    QVariant data;

    QAxObject *cell = pWorksheet->querySubObject("Cells(int,int)",row,column);//��ȡ��Ԫ�����
    if ( cell )
    {
        data = cell->dynamicCall("Value2()");
    }

    return data;
}

/**
  *@brief �޸�ָ����Ԫ�������
  *@param row : ��Ԫ����к�
  *@param column : ��Ԫ��ָ�����к�
  *@param data : ��Ԫ��Ҫ�޸�Ϊ��������
  *@return �޸��Ƿ�ɹ� true : �ɹ�
  *                   false: ʧ��
  */
bool ExcelEngine::SetCellData(UINT row, UINT column, QVariant data)
{
    bool op = false;

    QAxObject *cell = pWorksheet->querySubObject("Cells(int,int)",row,column);//��ȡ��Ԫ�����
    if ( cell )
    {
        QString strData = data.toString(); //excel ��Ȼֻ�ܲ����ַ��������ͣ��������޷�����
        cell->dynamicCall("SetValue(const QVariant&)",strData); //�޸ĵ�Ԫ�������
        op = true;
    }
    else
    {
        op = false;
    }

    return op;
}

/**
  *@brief ��ճ�����֮�������
  */
void ExcelEngine::Clear()
{
    sXlsFile     = "";
    nRowCount    = 0;
    nColumnCount = 0;
    nStartRow    = 0;
    nStartColumn = 0;
}

/**
  *@brief �ж�excel�Ƿ��ѱ���
  *@return true : �Ѵ�
  *        false: δ��
  */
bool ExcelEngine::IsOpen()
{
    return bIsOpen;
}

/**
  *@brief �ж�excel COM�����Ƿ���óɹ���excel�Ƿ����
  *@return true : ����
  *        false: ������
  */
bool ExcelEngine::IsValid()
{
    return bIsValid;
}

/**
  *@brief ��ȡexcel������
  */
UINT ExcelEngine::GetRowCount()const
{
    return nRowCount;
}

/**
  *@brief ��ȡexcel������
  */
UINT ExcelEngine::GetColumnCount()const
{
    return nColumnCount;
}
QVariant ExcelEngine::castListListVariant2Variant(int startLineNum)
{
    QList<QList<QVariant> > res_1 = res;
    QVariantList vars_1;
    const int rows = res_1.size();
    for(int i=startLineNum;i<rows;++i)
    {
        vars_1.append(QVariant(res_1[i]));
    }
    return QVariant(vars_1);
}
QList<QList<QVariant>> ExcelEngine::getRes()
{
    return res;
}
QVariant ExcelEngine::getVar()
{
    return var;
}
