using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
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

                for (int k = 0; k < wk.NumberOfSheets; k++)
                {
                    var hs = wk.GetSheetAt(k);
                    var hr = hs.GetRow(0); //row

                    int i = 0;
                    hr = hs.GetRow(i); //column
                    for (int j = hr.FirstCellNum; j < hr.LastCellNum; j++)
                    {
                        Console.WriteLine("run:{0}",j);
                        if (hr.GetCell(j) != null)
                        {
                            Console.Write("({0},{1}) = {2} ; ", i, j, hr.GetCell(j).ToString());
                        }
                    }
                }
                Console.WriteLine("finish");

                /*
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
                */

                wk = null; //全部Sheet讀完關閉Excel
                fs.Close();

                foreach (var item in modelList)
                {
                    Console.Write(item.setProductNumber);
                }
                Console.Read();
            }
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
            excel.ExcelToList();
        }
    }
}
