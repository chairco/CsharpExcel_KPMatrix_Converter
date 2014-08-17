using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Reflection;
using System.Data;
using System.Collections;
using System.Collections.Specialized;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using System.IO;
using NPOI.HSSF.Util;
using NPOI.HSSF.UserModel;

/*
透過NPOI來讀取EXCEL,要先下載NPOI DLL去參考,說明如下:
NPOI.DLL：NPOI 核心函式庫。
NPOI.DDF.DLL：NPOI 繪圖區讀寫函式庫。
NPOI.HPSF.DLL：NPOI 文件摘要資訊讀寫函式庫。
NPOI.HSSF.DLL：NPOI Excel BIFF 檔案讀寫函式庫。
NPOI.Util.DLL：NPOI 工具函式庫。
NPOI.POIFS.DLL：NPOI OLE 格式存取函式庫。
ICSharpCode.SharpZipLib.DLL：檔案壓縮函式庫。
*/

namespace BOM_Matrix_cmd
{
    class _EXCEL
    {
        private string fileName = null; // Data name
        private string Datasource = null;
        
        public List<string> configList = new List<string>(); //存config
        public List<string> complist = new List<string>(); //存component
        public Hashtable db = new Hashtable(); //根據config存component
        public Hashtable db_comp = new Hashtable();
        public Hashtable db_config = new Hashtable();

        public _EXCEL(string fileName)
        {
            this.fileName = fileName;
            this.Datasource = fileName;
        }

        //將config 29~31的內容存入DataTable
        public DataTable ExcelToConfig()
        {
            DataTable dt = new DataTable();
            IWorkbook wk = null;
            using (FileStream fs = new FileStream(fileName, FileMode.Open, FileAccess.ReadWrite))
            {
                if (Datasource.Contains(".xlsx")) //2007
                {
                    wk = new XSSFWorkbook(fs);
                }
                else //2003
                {
                    wk = new HSSFWorkbook(fs);
                }

                Console.WriteLine("Config寫入記憶體");
                //因為只有一個Sheet
                for (int k = 0; k < wk.NumberOfSheets; k++)
                {
                    ISheet st = wk.GetSheetAt(k);
                    IRow row = null;
                    String value = null;
                    DataRow dr = null;

                    //開始讀表格,一列開始讀(row,x軸)再讀行(其實是讀每個cell)
                    for (int i = st.FirstRowNum + 1; i <= 30; i++)
                    {
                        dr = dt.NewRow();
                        row = st.GetRow(i);

                        if (row != null)
                        {
                            for (int j = 29; j < row.LastCellNum; j++)
                            {
                                if (row.GetCell(j) != null) //解決跨行cell是空值問題
                                {
                                    value = row.GetCell(j).ToString().ToUpper();
                                    //Console.Write("({0},{1})={2}, ",i,j,value);
                                    //Console.Write(value + ",");
                                    if(i==2 && j>=30)
                                    {
                                        //Console.WriteLine(value);
                                        configList.Add(value);
                                    }
                                }
                            }
                            //Console.Write("\n");
                        }
                    }
                }
                wk = null; //全部Sheet讀完關閉Excel
                fs.Close();
            }
            return dt;
        }
        
        public DataTable ExcelToDataTable()
        {
            DataTable dt = new DataTable();
            IWorkbook wk = null;
            Boolean column_flag = false;
            Boolean dt_flag = false;
            String Item_value = null;
            String Comp_value = null;
            
            using (FileStream fs = new FileStream(fileName, FileMode.Open, FileAccess.ReadWrite))
            {
                if (Datasource.Contains(".xlsx")) //2007
                {
                    wk = new XSSFWorkbook(fs);
                }
                else //2003
                {
                    wk = new HSSFWorkbook(fs);
                }
                //只有一個Sheet,k=0
                for (int k = 0; k < wk.NumberOfSheets; k++)
                {
                    ISheet st = wk.GetSheetAt(k);
                    IRow row = null;
                    String value = null;
                    DataRow dr = null;
                    CellType ct = CellType.Blank;

                    Console.WriteLine("Item寫入記憶體\n");
                    //開始讀表格,一列開始讀(row,x軸)再讀行(其實是讀每個cell)
                    for (int i = st.FirstRowNum; i <= st.LastRowNum; i++)
                    {
                        dr = dt.NewRow();
                        row = st.GetRow(i);

                        if (row != null)
                        {
                            for (int j = row.FirstCellNum; j < row.LastCellNum; j++)
                            {
                                if (row.GetCell(j) != null) //解決跨行cell是空值問題
                                {
                                    value = row.GetCell(j).ToString().ToUpper();

                                    //只抓從ITEM欄位開始的資料
                                    if (value == "ITEM")
                                    {
                                        //Console.WriteLine("開始抓Column");
                                        column_flag = true; //用來判斷標頭
                                        dt_flag = true; //用來判斷開始抓入dt
                                    }
                                    //因為ITEM,COPMPNENT是一個跨行CELL將她獨立出來,加在每一個dt
                                    if (dt_flag && row.GetCell(0).IsMergedCell)　//橫向只需要抓一遍就跳出
                                    {
                                        Item_value = row.GetCell(0).ToString().ToUpper();
                                        dr[j] = Item_value;
                                        break;
                                    }
                                    if (dt_flag && row.GetCell(2).CellType.ToString() != "Blank")　//直向要判斷是否有新的值
                                    {
                                        Comp_value = row.GetCell(2).ToString().ToUpper();
                                        if (!complist.Contains(Comp_value) && Comp_value != "COMPONENT") complist.Add(Comp_value);
                                    }

                                    //標題
                                    if (column_flag && dt_flag)
                                    {
                                        //Console.WriteLine("Column座標 ({0},{1}) = {2}", i, j, value);
                                        dt.Columns.Add(row.GetCell(j).StringCellValue.Trim());
                                        dr[j] = row.GetCell(j);
                                    }
                                    //內容
                                    if (!column_flag && dt_flag) //可能會有公式欄位,需要改個寫法
                                    { 
                                        //Console.WriteLine("Row座標 ({0},{1}) = {2}", i, j, value);
                                        if (j == 0)
                                        {
                                            dr[j] = Item_value;
                                        }
                                        else if (j == 2)
                                        {
                                            dr[j] = Comp_value;
                                        }
                                        else
                                        {
                                            //dr[j] = value;//原本如果不做公式轉換，會有錯誤

                                            ct = row.GetCell(j).CellType;
                                            //如果此欄位格式為公式 則去取得CachedFormulaResultType
                                            if (ct == CellType.Formula)
                                            {
                                                ct = row.GetCell(j).CachedFormulaResultType;
                                            }
                                            if (ct == CellType.Numeric)
                                            {
                                                dr[j] = row.GetCell(j).NumericCellValue;
                                            }
                                            else
                                            {
                                                dr[j] = row.GetCell(j).ToString().Replace("$", "");
                                            }
                                        }
                                    }
                                }
                            }
                            if (dt_flag) dt.Rows.Add(dr);
                            column_flag = false;
                        }
                    }
                }
                wk = null; //全部Sheet讀完關閉Excel
                fs.Close();
            }
            return dt;
        }

        public void DataTableToHashTable(DataTable dt)
        {
            int ItemList_row = 0;
            DataRow dr2 = null;

            //根據存在list config一個一個讀出所有數據並存在HashTable
            foreach(string config in configList)
            {
                DataTable dt2;
                dt2 = dt.Clone(); //dt2建立一個和dt一樣的schmea,不包含rows data只包含columns data

                //找出config row位置
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    //Console.Write("{0} = {1},",i,dt.Columns[i].ColumnName);
                    if (dt.Columns[i].ColumnName == config) //選擇config
                    {
                        //Console.WriteLine("{0} site= {1}", dt.Columns[i].ColumnName, i);
                        ItemList_row = i;
                    }
                }

                //根據config位置找出row並儲存到dt2
                for (int i = 0; i < dt.Rows.Count; i++)
                {

                    //尋找X位置,並儲存那一列
                    if (dt.Rows[i][ItemList_row].ToString().ToUpper() == "X")
                    {
                        dr2 = dt2.NewRow();
                        for (int j = 0; j < dt.Columns.Count; j++)
                        {
                            dr2[j] = dt.Rows[i][j];
                        }
                        dt2.Rows.Add(dr2);
                    }
                }
                db.Add(config, dt2); //建立CONFIG和對應零件到HashTable
            }
            //DisplayDT(dt2); //列印dt2值

            //DisplayHash(db);
        }
        
        public void DataTableToExcel()
        {

            string[] item_out = new string[] {"VENDOR", "CONFIG", "NOTES" };
            string component = "COMPONENT";
            DataRow dr = null;

            //建立一個columns=vendor,config,notes的欄位
            DataTable dt_items = new DataTable();
            dt_items.Columns.Add(new DataColumn("vendor"));
            dt_items.Columns.Add(new DataColumn("config"));
            dt_items.Columns.Add(new DataColumn("notes"));
            int item_site = 0;

            foreach (DictionaryEntry one in db) //hashtable read by config
            {
                Console.WriteLine("key = " + one.Key);
                DataTable value = (DataTable)one.Value;

                for (int k = 0; k < value.Columns.Count; k++) //search columns(欄位標題)
                {
                    if (value.Columns[k].ColumnName.ToUpper() == component) //為component
                    {
                        item_site = k;
                        for (int i = 0; i < value.Rows.Count; i++) //讀這一列
                        {
                            int m = 0;
                            Console.WriteLine(value.Rows[i][k].ToString() + ","); //從這裡開始，建立一個dt根據每個compnonet去存
                            DataTable dt_item;
                            dt_item = dt_items.Clone();

                            //建立一個新行-->對應dt_item.Rows.Add(dr)
                            dr = dt_item.NewRow();
                            foreach (string items in item_out) //搜尋要找的欄位標題
                            {   
                                for (int n = 0; n < value.Columns.Count; n++)
                                {
                                    if (value.Columns[n].ColumnName.ToUpper() == items.ToUpper())
                                    {       
                                        //int _item_site = n;
                                        Console.WriteLine("m="+m+","+value.Rows[i][n].ToString());
                                        dr[m] = value.Rows[i][n];
                                        m += 1;
                                    }
                                }
                            }
                            dt_item.Rows.Add(dr);
                            DisplayDT(dt_item);
                            Console.WriteLine("********************");
                        }
                    }
                }
                Console.WriteLine("--------------------");
            }
            Console.Read();
        }

        public static int DisplayDT(DataTable dt)
        {
            //for (int i = 0; i < dt.Columns.Count; i++)
            //{
            //    Console.Write(dt.Columns[i].ColumnName+", ");
            //}
            //Console.Write("\n\n");
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    Console.WriteLine("({0},{1}) = {2}",i,j,dt.Rows[i][j].ToString());
                }
            }
            //Console.WriteLine("\n"+dt.Rows.Count);
            //Console.WriteLine(dt.Columns.Count);
            //Console.Read();
            return 0;
        }

        public static int DisplayHash(Hashtable ht)
        {
            foreach (DictionaryEntry one in ht)
            {
                Console.WriteLine("key = " + one.Key);
                DataTable value = (DataTable)one.Value;
                for (int i = 0; i < value.Rows.Count; i++)
                {
                    for (int j = 0; j < value.Columns.Count; j++)
                    {
                        Console.Write(value.Rows[i][j].ToString() + ",");
                    }
                    Console.Write("\n\n");
                }
            }
            Console.Read();
            return 0;
        }
    }

    class Program
    {
        //main program
        static void Main(string[] args)
        {
            DataTable dt = new DataTable();
            
            //read excel path
            string file = "C:\\Panda.xlsx";
            _EXCEL excel = new _EXCEL(file);

            //read excel config data to Datatable and arrayList
            excel.ExcelToConfig();

            //read excel item data to Datatable
            dt = excel.ExcelToDataTable();
            excel.DataTableToHashTable(dt);
            excel.DataTableToExcel();

            //excel.DisplayDataTable(excel.ExcelToDataTable());
            //excel.DisplayDataTable(excel.GetDataTableFromExcelFile());
        }
    }
}
