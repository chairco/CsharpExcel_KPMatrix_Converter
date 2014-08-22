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


namespace BOM_Matrix_cmd
{
    class Class1
    {
        public DataTable GetDataTableFromExcelFile(string fileName)
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

        /*
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

                //因為只有一個Sheet
                for (int k = 0; k < wk.NumberOfSheets; k++)
                {
                    ISheet st = wk.GetSheetAt(k);
                    IRow row = null;
                    String value = null;
                    DataRow dr = null;

                    Console.WriteLine("開始將Config寫入記憶體");
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
                                    if (i == 2 && j >= 30)
                                    {
                                        if (value != "CONFIGS") configList.Add(value);
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
        */
    }
}

//原本印出程式(有bug)
//foreach (DictionaryEntry config_key in db_config) //CONFIG
//{
//    Console.WriteLine("CONFIG = " + config_key.Key.ToString());

//    foreach (DictionaryEntry value in (Hashtable)config_key.Value) //COMP
//    {
//        Console.WriteLine("COMP = "+value.Key.ToString());

//        if (!ck_comp2)
//        {
//            sheet1.CreateRow(x + 1); //第x+1行
//            sheet1.GetRow(x + 1).CreateCell(y + 1).SetCellValue("Config Name");
//        }
//        if (!ck_comp3)
//        {
//            sheet1.GetRow(x + 1).CreateCell(y + 2).SetCellValue(config_key.Key.ToString());
//            //Console.WriteLine("{0},{1} = {2}", x + 1, y + 2, config_key.Key.ToString());
//        }

//        DataTable values = (DataTable)value.Value;
//        data = null;
//        for (int i = 0; i < values.Rows.Count; i++)
//        {
//            for (int j = 0; j < values.Columns.Count; j++)
//            {
//                //Console.Write(values.Rows[i][j].ToString() + ",");
//                data += values.Rows[i][j].ToString()+";";
//            }
//        }
//        Console.WriteLine(data);

//        if (!ck_comp)
//        {
//            sheet1.CreateRow(x + 2);
//            sheet1.GetRow(x + 2).CreateCell(y + 1).SetCellValue(value.Key.ToString()); //寫入EXCEL //第n+2行
//            temp = x + 2;
//        }

//        if (temp < x + 2)
//        {
//            sheet1.CreateRow(x + 2);
//        }

//        sheet1.GetRow(x + 2).CreateCell(y + 2).SetCellValue(data);//寫入數值
//        //Console.WriteLine("({0},{1})", x + 2, y + 2);

//        x += 1;
//        ck_comp2 = true;
//        ck_comp3 = true;
//    }
//    y += 1;
//    x = 0;
//    ck_comp = true;
//    ck_comp3 = false;
//    Console.WriteLine("----------------------");
//}