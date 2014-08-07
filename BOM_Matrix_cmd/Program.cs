using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using System.IO;
using NPOI.HSSF.Util;
using NPOI.HSSF.UserModel;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Reflection;
using System.Data;
using System.Data.OleDb;
using System.Collections;
using System.Collections.Generic; //list1
using System.Collections.Specialized;

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
    public class SheetModifyProductList
    {
        public string setProductNumber { get; set; }
    }

    class EXCEL
    {
        private string fileName = null; // Data name
        private string Datasource = null;

        public EXCEL(string fileName)
        {
            this.fileName = fileName;
            this.Datasource = fileName;
        }

        public void ExcelToList()
        {
            IWorkbook wk;
            ISheet st;
            List<SheetModifyProductList> modelList = new List<SheetModifyProductList>(); //add

            using (FileStream fs = new FileStream(Datasource, FileMode.Open, FileAccess.Read))
            {
                if (Datasource.Contains(".xlsx")) //2007
                {
                    wk = new XSSFWorkbook(fs);
                    st = (XSSFSheet)wk.GetSheetAt(0);
                }
                else //2003
                {
                    wk = new HSSFWorkbook(fs);
                    st = (HSSFSheet)wk.GetSheetAt(0);
                }

                /*
                //每個Sheet都會讀
                for (int k = 0; k < wk.NumberOfSheets; k++)
                {
                    //設定hs為某一個Sheet
                    var hs = wk.GetSheetAt(k);
                    
                    //設定Row(X軸),一開始從0開始
                    var hr = hs.GetRow(0);
                    int j = 0;
                    for (int i = hs.FirstRowNum; i <= hs.LastRowNum; i++)
                    {
                        if (hr.GetCell(j) != null)
                        {
                            Console.WriteLine("({0},{1}) = {2} ; ", i, j, hr.GetCell(i).ToString());
                        }
                    }
                    Console.WriteLine("finish");
                }
                */
                for (int k = 0; k < wk.NumberOfSheets; k++) //讀出sheetname
                {
                    var hs = wk.GetSheetAt(k); //sheet
                    string sheetname = hs.SheetName.ToString(); //sheet's name
                    Console.WriteLine("Sheet name: {0}.{1}", k, sheetname);
                    if (sheetname != "N61 FF") continue;

                    var hr = hs.GetRow(0); //row
                    for (int i = hs.FirstRowNum; i <= hs.LastRowNum; i++)
                    {
                        hr = hs.GetRow(i); //column
                        for (int j = hr.FirstCellNum; j < hr.LastCellNum; j++)
                        {
                            SheetModifyProductList model = new SheetModifyProductList(); //add

                            if (i == 0 && hr.GetCell(j) != null)
                            {
                                //Console.Write("({0},{1}) = {2} ; ", i, j, hr.GetCell(j).ToString());
                                model.setProductNumber = hr.GetCell(j).ToString() + ";";
                            }
                            modelList.Add(model);
                        }
                        //Console.WriteLine("\t\n");
                    }
                }

                wk = null; //全部Sheet讀完關閉Excel
                fs.Close();

                foreach (var item in modelList)
                {
                    Console.Write(item.setProductNumber);
                }
                Console.Read();
            }
        }

        public void ReadExcel()
        {
            IWorkbook WK;
            using (FileStream fs = new FileStream(fileName, FileMode.Open, FileAccess.ReadWrite))
            {
                if (Datasource.Contains(".xlsx")) //2007
                {
                    WK = new XSSFWorkbook(fs);
                }
                else //2003
                {
                    WK = new HSSFWorkbook(fs);
                }
                
                ISheet sheet = WK.GetSheet("FATP");

                for (int column = 0; column <= sheet.LastRowNum; column++)
                {
                    if (column >= 0 && column <= 2)
                    {
                        if (sheet.GetRow(column).GetCell(0) != null && sheet.GetRow(column).GetCell(0).ToString() != "")
                        {
                            Console.WriteLine("Row {0} = {1}", column, sheet.GetRow(column).GetCell(0).StringCellValue);
                        }
                    }
                }
            }
            Console.Read();
        }

        //設定ConfigHashTable
        public class ConfigHashTable
        {
            public string setConfigAsix { get; set; }
        }

        public Hashtable ConfigName = new Hashtable();

        public void TestExcel()
        { 
            IWorkbook wk = null;
            ISheet st;
            bool flag = false; //用來判斷是否讀到configs
            
            //New confighashtable
            HashSet<ConfigHashTable> configlist = new HashSet<ConfigHashTable>();

            using (FileStream fs = new FileStream(fileName, FileMode.Open, FileAccess.ReadWrite))
            {
                //new confighash為SetConfigAsix
                ConfigHashTable ConfigHash = new ConfigHashTable();
                
                if (Datasource.Contains(".xlsx")) //2007
                {
                    wk = new XSSFWorkbook(fs);
                }
                else //2003
                {
                    wk = new HSSFWorkbook(fs);
                }

                for (int k = 0; k < wk.NumberOfSheets; k++)
                {
                    st = wk.GetSheetAt(k);
                    IRow row = st.GetRow(k); //當前行數

                    //一列讀x軸
                    for (int i = 0; i <= st.LastRowNum; i++)
                    {
                        row = st.GetRow(i);
                        if (row != null)
                        {
                            //一列內(x軸)的cell讀出來
                            for (int j = 0; j < row.LastCellNum; j++)
                            {
                                if (row.GetCell(j) != null)
                                {
                                    string value = row.GetCell(j).ToString();
                                    if (value == "Configs")
                                    {
                                        flag = true;
                                    }
                                    if (flag == true && value != "Configs")
                                    {
                                        Console.WriteLine("Config={0}, X={1},",value.ToString(),j);
                                        ConfigName.Add(value.ToString(), j);
                                    }
                                }
                            }
                            flag = false;
                        }
                    }
                }
            wk = null; //全部Sheet讀完關閉Excel
            fs.Close();
            }
            Console.WriteLine("finish");
            Console.Read();
            

            //foreach (DictionaryEntry one in ConfigName)
            //{
            //    String infoString = "";
            //    object savedDetail = one.Value;
            //    foreach (DictionaryEntry detail in (Hashtable)savedDetail)
            //    {
            //        infoString = infoString + detail.Key.ToString() + " => " + detail.Value.ToString() + "\r\n";
            //        Console.WriteLine("索引鍵:{0},值:{1}", detail.Key, detail.Value);
            //    }
            //}
        }
    }

    class sample_test
    {
        public void test()
        {
            //建立Excel 2003檔案
            IWorkbook wb = new HSSFWorkbook();
            ISheet ws = wb.CreateSheet("Class");

            ////建立Excel 2007檔案
            //IWorkbook wb = new XSSFWorkbook();
            //ISheet ws = wb.CreateSheet("Class");

            ws.CreateRow(0);//第一行為欄位名稱
            ws.GetRow(0).CreateCell(0).SetCellValue("name");
            ws.GetRow(0).CreateCell(1).SetCellValue("score");
            
            ws.CreateRow(1);//第二行之後為資料
            ws.GetRow(1).CreateCell(0).SetCellValue("abey");
            ws.GetRow(1).CreateCell(1).SetCellValue(85);
            ws.CreateRow(2);
            ws.GetRow(2).CreateCell(0).SetCellValue("tina");
            ws.GetRow(2).CreateCell(1).SetCellValue(82);
            ws.CreateRow(3);
            ws.GetRow(3).CreateCell(0).SetCellValue("boi");
            ws.GetRow(3).CreateCell(1).SetCellValue(84);
            ws.CreateRow(4);
            ws.GetRow(4).CreateCell(0).SetCellValue("hebe");
            ws.GetRow(4).CreateCell(1).SetCellValue(86);
            ws.CreateRow(5);
            ws.GetRow(5).CreateCell(0).SetCellValue("paul");
            ws.GetRow(5).CreateCell(1).SetCellValue(82);
            
            FileStream file = new FileStream(@"c:\npoi.xls", FileMode.Create);//產生檔案
            wb.Write(file);
            file.Close();
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            string file = "C:\\Panda.xlsx";
            EXCEL excel = new EXCEL(file);
            excel.TestExcel();
            //excel.ReadExcel();
            //excel.ExcelToList();
        }
    }
}
