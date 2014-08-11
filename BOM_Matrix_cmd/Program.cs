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

        public Hashtable ht = new Hashtable();

        public _EXCEL(string fileName)
        {
            this.fileName = fileName;
            this.Datasource = fileName;
        }

        public DataTable ExcelToDataTable()
        {
            DataTable dt = new DataTable();
            List<string> ItemList = new List<string>();

            IWorkbook wk = null;
            ISheet st;
            Boolean column_flag = false;
            Boolean dt_flag = false;
            Boolean comp_flag = true;
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

                //依照順序讀每個Sheet(k)
                for (int k = 0; k < wk.NumberOfSheets; k++)
                {
                    st = wk.GetSheetAt(k);
                    IRow row = st.GetRow(k); //當前行數
                    DataRow dr = null;
                    String value = null;

                    //開始讀表格,一列開始讀(row,x軸)再讀行(其實是讀每個cell)
                    for (int i = 0; i <= st.LastRowNum; i++)
                    {
                        dr = dt.NewRow();
                        row = st.GetRow(i);
                        if (row != null)
                        {
                            for (int j = 0; j < row.LastCellNum; j++)
                            {
                                if (row.GetCell(j) != null) //解決跨行cell是空值問題
                                {
                                    value = row.GetCell(j).ToString().ToUpper();

                                    //只抓從ITEM欄位開始的資料
                                    if (value == "ITEM")
                                    {
                                        //Console.WriteLine("開始抓Column");
                                        column_flag = true;
                                        dt_flag = true;
                                    }
                                    
                                    //因為ITEM,COPMPNENT是一個跨行CELL將她獨立出來,加在每一個dt
                                    if (dt_flag && row.GetCell(0).IsMergedCell)
                                    {
                                        Item_value = row.GetCell(0).ToString().ToUpper();
                                        //Console.WriteLine(Item_value);
                                        break;
                                    }
                                    if (dt_flag && row.GetCell(2).IsMergedCell && comp_flag)
                                    {
                                        Comp_value = row.GetCell(2).ToString().ToUpper();
                                        //Console.WriteLine(Comp_value);
                                        comp_flag = false;
                                    }

                                    //標題
                                    if (column_flag && dt_flag)
                                    {
                                        //Console.WriteLine("Column座標 ({0},{1}) = {2}", i, j, value);
                                        dt.Columns.Add(row.GetCell(j).StringCellValue.Trim());
                                    }
                                    
                                    //內容
                                    if (!column_flag && dt_flag) //可能會有公式欄位,需要改個寫法
                                    { 
                                        //Console.WriteLine("Row座標 ({0},{1}) = {2}", i, j, value);
                                        if (j == 0)
                                        {
                                            dt.Rows.Add(Item_value);
                                        }
                                        else if (j == 2)
                                        {
                                            dt.Rows.Add(Comp_value);
                                        }
                                        else
                                        {
                                            dt.Rows.Add(value);
                                        }
                                    }
                                }
                            }
                            column_flag = false;
                        }
                    }
                }
                wk = null; //全部Sheet讀完關閉Excel
                fs.Close();
            }
            return dt;
        }

        public void DisplayHashTable(Hashtable data)
        {
            foreach (DictionaryEntry Table in data)
            {
                Console.WriteLine("索引鍵:{0},值:{1}", Table.Key, Table.Value);
            }
        }

        public void DisplayDataTable(DataTable dt)
        {
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                Console.WriteLine("value{0} = {1}",i,dt.Rows[i][0].ToString());
            }

            Console.WriteLine("\n" + dt.Rows.Count);
            Console.WriteLine(dt.Columns.Count);
            
            Console.Read();
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            string file = "C:\\Panda.xlsx";
            _EXCEL excel = new _EXCEL(file);
            excel.DisplayDataTable(excel.ExcelToDataTable());
        }
    }
}
