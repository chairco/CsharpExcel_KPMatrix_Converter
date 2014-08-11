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
using System.Collections;
using System.Collections.Specialized;


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

        public Hashtable ItemList = new Hashtable();

        public DataTable TestExcel()
        {
            IWorkbook wk = null;
            ISheet st;
            DataTable dt = new DataTable();
            bool flag = false; //用來判斷是否讀到configs

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

                for (int k = 0; k < wk.NumberOfSheets; k++)
                {
                    st = wk.GetSheetAt(k);
                    IRow row = st.GetRow(k); //當前行數

                    ////取得Merge大小
                    //for (int r = 0; r < st.NumMergedRegions; r++)
                    //{
                    //    Console.WriteLine(st.GetMergedRegion(r));
                    //}

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
                                    string value = row.GetCell(j).ToString().ToUpper();

                                    //判斷是否有merge cell
                                    if (value == "MLB")
                                    {
                                        Console.WriteLine("座標=({0},{1})", i, j);
                                        if (row.GetCell(j).IsMergedCell)
                                        {
                                            Console.WriteLine(row.GetCell(j).IsMergedCell.ToString());
                                        }
                                    }
                                    if (value == "ITEM")
                                    {
                                        flag = true;
                                    }
                                    if (flag == true)
                                    {
                                        Console.WriteLine("Item= {0}, X= {1},", value.ToString(), j);
                                        if (!ItemList.ContainsKey("QOH"))
                                        {
                                            ItemList.Add(value.ToString(), j);
                                            dt.Columns.Add(row.GetCell(j).StringCellValue.Trim());
                                        }
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
            //Display(ConfigName);
            //Display(ItemList);
            Console.Read();
            return dt;
        }

        public void Display(Hashtable data)
        {
            foreach (DictionaryEntry Table in data)
            {
                Console.WriteLine("索引鍵:{0},值:{1}", Table.Key, Table.Value);
            }
        }

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
                        if (dt.Columns.Contains("QOH"))
                        {
                            dt.Columns.Add(headerRow.GetCell(i).StringCellValue.Trim());
                        }
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
            Console.WriteLine("finish");
            return dt;
        }
    }
}
