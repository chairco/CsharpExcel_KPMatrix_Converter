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
    class npoi
    {
        private DataTable dt = new DataTable();
        public string fileName { get; set; }

        public DataTable ExcelToDataTable()
        {
            string Datasource = fileName;
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
                //只有一個Sheet,k=0
                for (int k = 0; k < wk.NumberOfSheets; k++)
                {
                    ISheet st = wk.GetSheetAt(k);
                    IRow row = null;
                    DataRow dr = null;
                    CellType ct = CellType.Blank;
               
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
                        dt.Rows.Add(dr);
                    }
                }
                wk = null; //全部Sheet讀完關閉Excel
                fs.Close();
            }
            return dt;
        }
    }
}
