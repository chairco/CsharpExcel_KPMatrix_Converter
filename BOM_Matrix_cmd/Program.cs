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
    public class _EXCEL
    {        
        private string fileName = null; // Data name
        private string Datasource = null;

        public List<string> configList = new List<string>(); //存config
        public List<string> complist = new List<string>(); //存component
        
        public Hashtable db = new Hashtable(); //根據config存component
        public Hashtable db_comp = new Hashtable(); //存component+整理過的datatable
        public Hashtable db_config = new Hashtable(); //存confgi+db_comp

        public _EXCEL(string fileName)
        {
            this.fileName = fileName;
            this.Datasource = fileName;
        }

        public static int DisplayDT(DataTable dt)
        {
            Console.WriteLine("DisPlay DataTable");
            //for (int i = 0; i < dt.Columns.Count; i++)
            //{
            //    Console.Write(dt.Columns[i].ColumnName + ", ");
            //}
            //Console.Write("\n\n");

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    //Console.WriteLine("({0},{1}) = {2}",i,j,dt.Rows[i][j].ToString());
                    Console.Write("({0},{1}) = {2}", i, j, dt.Rows[i][j].ToString() + ", ");
                }
                Console.WriteLine("\n------------");
            }
            Console.WriteLine("\n************");
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
                    Console.Write("\n");
                }
            }
            return 0;
        }
        
        public DataTable ExcelToDataTable(int comp_cell, bool mode, bool config)
        {
            DataTable dt = new DataTable();
            IWorkbook wk = null;
            
            Boolean column_flag = false;
            Boolean dt_flag = false;
            String Item_value = null;
            String Comp_value = null;
            int count = 2; //用來紀錄重複config key
            
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

                    if (!config)
                    {
                        Console.WriteLine("開始將Item寫入記憶體\n");
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
                                        if (dt_flag && row.GetCell(comp_cell).CellType.ToString() != "Blank")　//直向要判斷是否有新的值
                                        {
                                            Comp_value = row.GetCell(comp_cell).ToString().ToUpper();
                                            if (mode)
                                            {
                                                if (!complist.Contains(Comp_value) && Comp_value != "COMPONENT") complist.Add(Comp_value);
                                            }
                                            else if (!mode)
                                            {
                                                if (!complist.Contains(Comp_value) && Comp_value != "REF DES") complist.Add(Comp_value);
                                            }
                                        }

                                        //標題
                                        if (column_flag && dt_flag)
                                        {
                                            //Console.WriteLine("Column座標 ({0},{1}) = {2}", i, j, value);
                                            if (dt.Columns.Contains(row.GetCell(j).StringCellValue.Trim()))
                                            {
                                                string temp = row.GetCell(j).StringCellValue.Trim() + "_" + count;
                                                dt.Columns.Add(temp);
                                                count++;
                                            }
                                            else
                                            {
                                                dt.Columns.Add(row.GetCell(j).StringCellValue.Trim());
                                            }
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
                                            else if (j == comp_cell)
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
                    else if (config)
                    {
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
                }
                wk = null; //全部Sheet讀完關閉Excel
                fs.Close();
            }
            //DisplayDT(dt);
            return dt;
        }

        public Hashtable DataTableToHashTable(DataTable dt)
        {
            int ItemList_row = 0;
            DataRow dr2 = null;

            //根據存在list config一個一個讀出所有數據並存在HashTable
            foreach(string config in configList)
            {
                //Console.WriteLine(config);
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
                    //FATP尋找X位置,並儲存那一列;如果是MLB尋找是1位置,並儲存那一列
                    if (dt.Rows[i][ItemList_row].ToString().ToUpper() == "X" || dt.Rows[i][ItemList_row].Equals("1"))
                    {
                        dr2 = dt2.NewRow();
                        for (int j = 0; j < dt.Columns.Count; j++)
                        {
                            dr2[j] = dt.Rows[i][j];
                        }
                        dt2.Rows.Add(dr2);
                    }
                }
                //DisplayDT(dt2); //列印dt2值
                db.Add(config, dt2); //建立CONFIG和對應零件到HashTable
            }
            //DisplayHash(db);
            return db;
        }
        
        public Hashtable DataTableToExcel(Hashtable db2, bool mode)
        {
            string[] item_out = null;
            string component = null;
            int item_site = 0;

            DataRow dr = null;
            //建立一個columns=vendor,config,notes的欄位
            DataTable dt_items = new DataTable();

            if (mode)
            {
                item_out = new string[] { "VENDOR", "CONFIG", "NOTES"};
                component = "COMPONENT";
                dt_items.Columns.Add(new DataColumn("vendor"));
                dt_items.Columns.Add(new DataColumn("config"));
                dt_items.Columns.Add(new DataColumn("notes"));
            }
            if (!mode)
            {
                item_out = new string[] { "SIDE", "MFR PN", "VENDOR", "MC Comment" };
                component = "REF DES";
                dt_items.Columns.Add(new DataColumn("side"));
                dt_items.Columns.Add(new DataColumn("mfr pn"));
                dt_items.Columns.Add(new DataColumn("vendor"));
                dt_items.Columns.Add(new DataColumn("mc comment"));
            }

            foreach (DictionaryEntry one in db2) //hashtable read by config
            {
                string config_Key = one.Key.ToString();
                //Console.WriteLine("CONFIG = " + one.Key);
                DataTable value = (DataTable)one.Value;

                for (int k = 0; k < value.Columns.Count; k++) //search columns(欄位標題)
                {
                    if (value.Columns[k].ColumnName.ToUpper() == component.ToUpper()) //為component
                    {
                        item_site = k;
                        for (int i = 0; i < value.Rows.Count; i++) //讀這一列
                        {
                            int m = 0;
                            string comp_hash_key = value.Rows[i][k].ToString();
                            
                            //Console.WriteLine(value.Rows[i][k].ToString() + ","); //從這裡開始，建立一個dt根據每個compnonet去存
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
                                        //Console.WriteLine("m="+m+","+value.Rows[i][n].ToString());
                                        dr[m] = value.Rows[i][n];
                                        m += 1;
                                    }
                                }
                            }
                            dt_item.Rows.Add(dr);
                            if (!db_comp.ContainsKey(comp_hash_key)) db_comp.Add(comp_hash_key, dt_item);
                            //DisplayDT(dt_item);
                            //Console.WriteLine("********************");
                        }
                        //DisplayHash(db_comp);
                        //Console.WriteLine(db_comp.Count);
                        object db_comp_cp = db_comp.Clone(); //建立一個Hashtable來接,不然db_comp.clear()會把所有資料清掉call by value
                        db_config.Add(config_Key, db_comp_cp);
                        //Console.WriteLine("********************");
                    }
                    db_comp.Clear();
                }
                //Console.WriteLine("--------------------");
            }
            return db_config;
            //WriteExcel();
            //Console.Read();
        }

        public void WriteExcel(Hashtable db_config2, List<string> INI)
        {
            Boolean ck_comp = false, ck_comp2 = false ;
            String data = null;
            IWorkbook wk = null;
            int x = 0, y = 0; //(x,y)
            
            wk = new XSSFWorkbook();
            
            // 新增試算表
            //wk.CreateSheet("試算表 A"); 

            //有資料內容的試算表
            XSSFSheet sheet1 = (XSSFSheet)wk.CreateSheet("Sheet1");

            //根據預先儲存的component去尋找每一個config
            foreach (string list in INI)
            //foreach (string list in complist) //原本是抓全部
            {
                //Console.WriteLine(list);
                if (!ck_comp)
                {                   
                    sheet1.CreateRow(x + 1); //第x+1行
                    sheet1.GetRow(x + 1).CreateCell(y + 1).SetCellValue("Config Name");
                    ck_comp = true;
                }
                sheet1.CreateRow(x + 2);
                sheet1.GetRow(x + 2).CreateCell(y + 1).SetCellValue(list); //寫入EXCEL //第n+2行
                sheet1.AutoSizeColumn(y + 1);
                
                //開始一個一個去讀config
                foreach (DictionaryEntry config_key in db_config2) //CONFIG
                {
                    //Console.WriteLine("CONFIG = " + config_key.Key.ToString());
                    //寫入config欄位
                    if (!ck_comp2)
                    {
                        //Console.WriteLine(config_key.Key.ToString());
                        sheet1.GetRow(x + 1).CreateCell(y + 2).SetCellValue(config_key.Key.ToString());
                    }
                            
                    foreach (DictionaryEntry value in (Hashtable)config_key.Value) //COMP
                    {
                        if(value.Key.ToString().ToUpper() == list.ToString().ToUpper())
                        {
                            //Console.WriteLine("COMP = " + value.Key.ToString());
                            DataTable values = (DataTable)value.Value;
                            data = null;
                            for (int i = 0; i < values.Rows.Count; i++)
                            {
                                for (int j = 0; j < values.Columns.Count; j++)
                                {
                                    //Console.Write(values.Rows[i][j].ToString() + ",");
                                    if (value.Key.ToString() == "MLB" && j == 1)
                                    {
                                        Console.WriteLine("{0}, {1}", config_key.Key.ToString() , values.Rows[i][j].ToString());
                                        ReadINI(false);
                                    }
                                    data += values.Rows[i][j].ToString() + "\n";
                                }
                            }
                            //Console.WriteLine("value=" + data);
                            //Console.WriteLine("value site = ({0},{1})", x+2, y+2);

                            sheet1.GetRow(x + 2).CreateCell(y + 2).SetCellValue(data);//寫入數值(component內容)
                            if (data.ToString().Contains("\n"))
                            {
                                //將目前欄位的CellStyle設定為自動換行
                                XSSFCellStyle cs = (XSSFCellStyle)wk.CreateCellStyle();
                                cs.WrapText = true;
                                sheet1.GetRow(x + 2).GetCell(y + 2).CellStyle = cs;

                                //因為換行所以愈設幫他Row的高度變成兩倍
                                sheet1.GetRow(x + 2).HeightInPoints = 3 * sheet1.DefaultRowHeight / 18;
                                sheet1.AutoSizeColumn(y + 2);
                            }
                            y += 1;
                            break;
                        }
                    }
                }
                ck_comp2 = true;
                y = 0;
                x += 1;
            }

            try
            {
                FileStream file = new FileStream(@"KeyPart_output.xlsx", FileMode.Create);
                wk.Write(file);
                file.Close();
            }
            catch (Exception ex)
            {
                Console.WriteLine("儲存檔案出錯，檔案可能正在使用" + Environment.NewLine + ex.Message);
            }
        }
        
        public List<string> ReadINI(bool mode)
        {
            List<string> INI = new List<string>();
            string[] lines = null;
            if (mode)
            {
                lines = System.IO.File.ReadAllLines(@"FATP.txt");
            }
            else if (!mode)
            {
                lines = System.IO.File.ReadAllLines(@"MLB.txt");
            }

            ////Display the file contents by using a foreach loop.
            //System.Console.WriteLine("Contents of WriteLines2.txt = ");
            foreach (string line in lines)
            {
                // Use a tab to indent each line of the file.
                string[] str = line.Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);
                foreach (string str2 in str)
                {
                    Console.Write(str2+", ");
                    INI.Add(str2);
                }
                //_fatp.Add(new data_list() { FATP = line});
            }
            Console.Write("\n");
            //// Keep the console window open in debug mode.
            //Console.WriteLine("Press any key to exit.");
            //System.Console.ReadKey();
            return INI;
        }
    }

    class Program
    {
        //main program
        static void Main(string[] args)
        {
            bool mode = true; //1,false; 2,true (MLB or FATP)

            //read excel path
            string file = "C:\\PANDA.xlsx";
            _EXCEL excel = new _EXCEL(file);

            //read excel config data to Datatable and arrayList
            excel.ExcelToDataTable(1, mode, true);

            //read excel item data to Datatable
            Hashtable ht = excel.DataTableToExcel(excel.DataTableToHashTable(excel.ExcelToDataTable(2, mode, false)), mode);
            excel.WriteExcel(ht, excel.ReadINI(mode));

            Console.WriteLine("Press any key to exit.");
            System.Console.ReadKey();

            //excel.DisplayDataTable(excel.ExcelToDataTable());
            //excel.DisplayDataTable(excel.GetDataTableFromExcelFile()); 
        }

        /*
        static bool blNoLogo = false;
        static bool blNoClearScreen = false;
        static bool blParameterOK = false;
        static bool blResult = false;
        static int nErrorLevel = 255;

        //static bool blPause = false;
        static void Main(string[] args)
        {
            try
            {
                for (int i = 0; i <= args.GetUpperBound(0); i++)
                {
                    if (string.Compare(args[i], "-nl", true) == 0)
                    {
                        blNoLogo = true;
                        blNoClearScreen = true;
                        continue;
                    }
                }

                if (!blNoLogo)
                    Diags.Logo(blNoLogo, !blNoClearScreen);

                for (int i = 0; i <= args.GetUpperBound(0); i++)
                {
                    if (string.Compare(args[i], "-erv", true) == 0)
                    {
                        if (args.GetUpperBound(0) > i)
                            nErrorLevel = int.Parse(args[i + 1]);
                        continue;
                    }
                }

                for (int i = 0; i <= args.GetUpperBound(0); i++)
                {
                    if (string.Compare(args[i], "/?", true) == 0)
                    {
                        Diags.ReadMe(blNoLogo);
                        continue;
                    }
                    else if (string.Compare(args[i], "/C", true) == 0)
                    {
                        _main(args[i + 1].ToString());
                        blParameterOK = true;
                    }
                }
                if (!blParameterOK)
                    Diags.ReadMe(blNoLogo);
            }
            catch (Exception ex)
            {
                Console.Write("缺少目標檔案: ");
                Console.WriteLine(ex.Message);
            }
            finally
            {
                if (blParameterOK && blResult)
                    Diags.End(0, blNoLogo);
                else
                    Diags.End(nErrorLevel, blNoLogo);
            }
        }

        static void _main(string file)
        {
            bool mode = true; //1,false; 2,true (MLB or FATP)

            //read excel path
            ////string file = "C:\\PANDA5.xlsx";
            _EXCEL excel = new _EXCEL(file);

            excel.ReadINI(true);

            //read excel config data to Datatable and arrayList
            excel.ExcelToDataTable(2, mode, true);

            //read excel item data to Datatable
            excel.DataTableToHashTable(excel.ExcelToDataTable(2, mode, false));
            excel.DataTableToExcel(mode);
        }
        */
    }
}
