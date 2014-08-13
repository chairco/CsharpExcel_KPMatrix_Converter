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

        //存config
        public List<string> configList = new List<string>();

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
            Boolean comp_flag = false;
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

                //因為只有一個Sheet
                for (int k = 0; k < wk.NumberOfSheets; k++)
                {
                    ISheet st = wk.GetSheetAt(k);
                    IRow row = null;
                    String value = null;
                    DataRow dr = null;
                    CellType ct = CellType.Blank;

                    Console.WriteLine("\nItem寫入記憶體\n");

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
                                    if (dt_flag && row.GetCell(0).IsMergedCell)
                                    {
                                        Item_value = row.GetCell(0).ToString().ToUpper();
                                        //Console.WriteLine(Item_value);
                                        dr[j] = Item_value;
                                        comp_flag = false; //當item開始,compoment也要重新抓
                                        break;
                                    }
                                    if (dt_flag && row.GetCell(2).IsMergedCell && !comp_flag)
                                    {

                                        Comp_value = row.GetCell(2).ToString().ToUpper();
                                        //Console.Write(Comp_value + ", ");
                                        comp_flag = true;

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

        public void DisplayDataTable(DataTable dt, string config)
        {
            int ItemList_row = 0;
            List<int> ItemList_col = new List<int>();
            DataRow dr2 = null;
            DataTable dt2;
            dt2 = dt.Clone(); //dt2建立一個和dt一樣的schmea,不包含rows只包含columns

            //找出config row位置
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                //Console.Write("{0} = {1},",i,dt.Columns[i].ColumnName);
                if (dt.Columns[i].ColumnName == config) //選擇config
                {
                    Console.WriteLine("{0} site= {1}", dt.Columns[i].ColumnName, i);
                    ItemList_row = i;
                }
            }
            Console.Write("\n");
            
            //印出和儲存config所有componemt data
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                
                //尋找X位置
                if (dt.Rows[i][ItemList_row].ToString().ToUpper() == "X")
                {
                    //印出row位置並儲存
                    //Console.WriteLine("({0},{1}) = {2}", i, ItemList_row, dt.Rows[i][ItemList_row].ToString()); //印出有ｘ的值
                    ItemList_col.Add(i);
                    dr2 = dt2.NewRow();
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        //Console.Write("({0},{1}) = {2}, ",i,j,dt.Rows[i][j]);
                        //Console.Write(dt.Rows[i][j]+", ");
                        dr2[j] = dt.Rows[i][j];
                    }
                    dt2.Rows.Add(dr2);
                    Console.Write("\n\n");
                }
            }
            Console.Write("\n\n");

            for (int i = 0; i < dt.Columns.Count; i++)
            {
                Console.Write(dt.Columns[i].ColumnName+", ");
            }

            Console.Write("\n\n");

            for (int i = 0; i < dt2.Rows.Count; i++)
            {
                for (int j = 0; j < dt2.Columns.Count; j++)
                {
                    Console.Write(dt2.Rows[i][j].ToString() + ", ");
                }
                Console.Write("\n\n");
            }

            //foreach (object prime in dt_list) // Loop through List with foreach
            //{
            //    Console.Write(prime+", ");
            //}

            Console.WriteLine("\n"+dt.Rows.Count);
            Console.WriteLine(dt.Columns.Count);
            Console.Read();
        }

        //sample
        public DataTable GetDataTableFromExcelFile()
        {
            FileStream fs = null;
            DataTable dt = new DataTable();
            try
            {
                IWorkbook wb = null;
                fs = File.Open(fileName, FileMode.Open, FileAccess.Read);
                switch (Path.GetExtension(fileName).ToUpper())
                {
                    case ".XLS":
                        {
                            wb = new HSSFWorkbook(fs);
                        }
                        break;
                    case ".XLSX":
                        {
                            wb = new XSSFWorkbook(fs);
                        }
                        break;
                }
                if (wb.NumberOfSheets > 0)
                {
                    ISheet sheet = wb.GetSheetAt(0);
                    IRow headerRow = sheet.GetRow(0);

                    //處理標題列
                    for (int i = headerRow.FirstCellNum; i < headerRow.LastCellNum; i++)
                    {
                        dt.Columns.Add(headerRow.GetCell(i).StringCellValue.Trim());
                    }

                    IRow row = null;
                    DataRow dr = null;
                    CellType ct = CellType.Blank;
                    
                    //標題列之後的資料
                    for (int i = sheet.FirstRowNum + 1; i <= sheet.LastRowNum; i++)
                    {
                        dr = dt.NewRow();
                        row = sheet.GetRow(i);
                        if (row == null) continue;
                        for (int j = row.FirstCellNum; j < row.LastCellNum; j++)
                        {
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
                        dt.Rows.Add(dr);
                    }
                }
                fs.Close();
            }
            finally
            {
                if (fs != null) fs.Dispose();
            }
            return dt;
        }
    }

    class Program
    {
        //main program
        static void Main(string[] args)
        {
            DataTable dt = new DataTable();
            
            //read excel path
            string file = "C:\\Panda_3.xlsx";
            _EXCEL excel = new _EXCEL(file);
            
            //read excel config data to Datatable and arrayList
            excel.ExcelToConfig();

            //read excel item data to Datatable
            dt = excel.ExcelToDataTable();
            excel.DisplayDataTable(dt,"G37");

            //excel.DisplayDataTable(excel.ExcelToDataTable());
            //excel.DisplayDataTable(excel.GetDataTableFromExcelFile());
        }
    }
}
