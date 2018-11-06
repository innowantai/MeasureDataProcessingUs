using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using System.Threading;
 
namespace ExcelClass
{ 
    public class EXCEL
    { 
        ///// public 公開屬性 : 提供外部存取之屬性，可取得Excel所有分頁名稱
        public List<string> sheets;
        ///// private 私有屬性 : 物件內部處理所使用之屬性
        private string FullPath;
        private Excel.Workbook workBook;
        private Excel.Application ExcelAPP;
        private Dictionary<string, string[,]> DicSheetsData = new Dictionary<string, string[,]>();

        /// <summary>
        /// 物件的建構函數
        /// </summary>
        /// <param name="fullPath"> Excel檔案路徑 </param>
        public EXCEL(string fullPath)
        {
            this.FullPath = fullPath;
            this.ExcelAPP = new Excel.Application();
            this.ExcelAPP.Visible = false;
            this.ExcelAPP.UserControl = true;
            this.ExcelAPP.DisplayAlerts = false;
            this.workBook = Is_Exist(this.FullPath);
            this.sheets = GetSheets();
        }


        //////// 以下為 public 公開方法(函數), 提供外部存取調用
        //////// 1.GetSheetsData      : 取得Excel所有Sheet的資料
        //////// 2.GetDataBySheetName : 取得Excel裡指定的Sheet資料
        //////// 3.Save_To            : 將資料儲存至其他Excel檔案中
        //////// 4.Save               : 將資料儲存至此開啟之Excel檔案中
        //////// 5.Close              : 關閉Excel程序

        /// <summary>
        /// 一次讀取所有分頁的資料，儲存至Dictionary格式變數中，其中keys為sheet名稱，values為excel資料
        /// </summary>
        /// <returns></returns>
        public Dictionary<string, string[,]> GetSheetsData()
        {
            foreach (string ss in sheets)
                this.DicSheetsData[ss] = GetDataBySheetName(1, 1, ss);

            return DicSheetsData;
        }

        /// <summary>
        /// 讀取指定sheet名稱的資料
        /// </summary>
        /// <param name="RowStartPo"> 指定列(row)的開始位置 </param>
        /// <param name="ColStartPo"> 指定欄(column)的開始位置 </param>
        /// <param name="SheetName"> 指定excel分頁的名稱 </param>
        /// <returns></returns>
        public string[,] GetDataBySheetName(int RowStartPo, int ColStartPo, string SheetName)
        {
             
            var SheetNum = sheets.IndexOf(SheetName) + 1;

            Excel.Worksheet ws = (Excel.Worksheet)this.workBook.Worksheets.get_Item(SheetNum);
            string[,] newData = ReadGetData(ws, RowStartPo, ColStartPo);

            System.Runtime.InteropServices.Marshal.ReleaseComObject(ws);
            return newData;
        }

        /// <summary>
        /// 儲存資料至指定的excel檔案 
        /// </summary>
        /// <param name="strPath"> excel檔案路徑 </param>
        /// <param name="sheetName"> 指定儲存之分頁名稱 </param>
        /// <param name="poRow"> 啟始列(row)的位置 </param>
        /// <param name="poCol"> 啟始欄(column)的位置 </param>
        /// <param name="Data"> 儲存資料 </param>
        public void Save_To(string strPath, string sheetName, int poRow, int poCol, string[,] Data)
        {
            bool fileExist = File.Exists(strPath);
            //// 若excel 不存在,創建
            if (!fileExist)
            { 
                Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
                excelApp.Visible = false;              //不顯示excel程式 
                excelApp.DisplayAlerts = false;        //设置是否显示警告窗体 
                Workbook book = excelApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
                Worksheet sheet = (Worksheet)book.Sheets[1];
                sheet.Name = sheetName;
                sheet.Activate();


                SaveFunction(strPath, Data, poRow, poCol, excelApp, sheet, book);
                 

            }
            else
            {
                //// excel 檔案已存在
                //// 创建Application
                Excel.Application excelApp = new Excel.Application();
                excelApp.DisplayAlerts = false;
                excelApp.Visible = false;
                excelApp.ScreenUpdating = false;

                Excel.Workbook book = excelApp.Workbooks.Open(strPath);
                int sheetCount = book.Sheets.Count;
                Worksheet ws = (Worksheet)book.Sheets[sheetCount];
                Worksheet sheet = null;
                //// 檢查sheets 是否已存在,
                bool exist = false;
                for (int i = 1; i < sheetCount + 1; i++)
                {
                    Worksheet indSheet = (Worksheet)book.Sheets[i];
                    if (sheetName == indSheet.Name)
                    {
                        exist = true;
                        break;
                    }
                }

                //// 若sheet 已存在,刪除重新建立,否則建立sheet
                if (exist)
                {
                    ws = (Worksheet)book.Sheets[sheetCount];
                    sheet = (Worksheet)book.Worksheets.Add(Type.Missing, ws, Type.Missing, Type.Missing);//建立一個新分頁 
                    Worksheet sheet2 = book.Sheets[sheetName];
                    sheet2.Delete();
                    sheet.Name = sheetName;
                }
                else
                {
                    sheet = (Worksheet)book.Worksheets.Add(Type.Missing, ws, Type.Missing, Type.Missing);//建立一個新分頁 
                    sheet.Name = sheetName;
                }



                SaveFunction(strPath, Data, poRow, poCol, excelApp, sheet, book);
                 
            }
        }

        /// <summary>
        /// 儲存資料至此開啟excel檔案
        /// </summary>
        /// <param name="sheetName"> 指定儲存之分頁名稱 </param>
        /// <param name="poRow"> 啟始列(row)的位置 </param>
        /// <param name="poCol"> 啟始欄(column)的位置 </param>
        /// <param name="Data"> 儲存資料 </param>
        public void Save(string sheetName, int poRow, int poCol, string[,] Data)
        {
            string strPath = this.FullPath;
            bool fileExist = File.Exists(strPath);
            //// 若excel 不存在,創建
            if (!fileExist)
            {
                FileInfo fi = new FileInfo(strPath);
                Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
                excelApp.Visible = false;              //不顯示excel程式 
                excelApp.DisplayAlerts = false;        //设置是否显示警告窗体 
                Workbook book = excelApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
                Worksheet sheet = (Worksheet)book.Sheets[1];
                sheet.Name = sheetName;
                sheet.Activate();


                SaveFunction(strPath, Data, poRow, poCol, excelApp, sheet, book);
                 

            }
            else
            {
                //// excel 檔案已存在
                //// 创建Application
                Excel.Application excelApp = new Excel.Application();
                excelApp.DisplayAlerts = false;
                excelApp.Visible = false;
                excelApp.ScreenUpdating = false;

                Excel.Workbook book = excelApp.Workbooks.Open(strPath);
                int sheetCount = book.Sheets.Count;
                Worksheet ws = (Worksheet)book.Sheets[sheetCount];
                Worksheet sheet = null;
                //// 檢查sheets 是否已存在,
                bool exist = false;
                for (int i = 1; i < sheetCount + 1; i++)
                {
                    Worksheet indSheet = (Worksheet)book.Sheets[i];
                    if (sheetName == indSheet.Name)
                    {
                        exist = true;
                        break;
                    }
                }

                //// 若sheet 已存在,刪除重新建立,否則建立sheet
                if (exist)
                {
                    ws = (Worksheet)book.Sheets[sheetCount];
                    sheet = (Worksheet)book.Worksheets.Add(Type.Missing, ws, Type.Missing, Type.Missing);//建立一個新分頁 
                    Worksheet sheet2 = book.Sheets[sheetName];
                    sheet2.Delete();
                    sheet.Name = sheetName;
                }
                else
                {
                    sheet = (Worksheet)book.Worksheets.Add(Type.Missing, ws, Type.Missing, Type.Missing);//建立一個新分頁 
                    sheet.Name = sheetName;
                } 

                SaveFunction(strPath, Data, poRow, poCol, excelApp, sheet, book);

            }
        }


        /// <summary>
        /// 關閉Excel物件程序
        /// </summary>
        public void close()
        {
            this.ExcelAPP.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(this.ExcelAPP);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(this.workBook);
            GC.Collect();
        }

        //////// 以下為 private 私有方法(函數), 物件內部處理所使用之方法(函數), 不提供外部存取

        /// <summary>
        /// 檢查指定的excel檔案是否存在，否則創立
        /// </summary>
        /// <param name="strPath"> excel檔案路徑 </param>
        /// <returns></returns>
        private Workbook Is_Exist(string strPath)
        { 
            bool fileExist = File.Exists(strPath);
            Workbook book;
            //// 若excel 不存在,創建
            if (!fileExist)
            { 
                Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
                excelApp.Visible = false;              //不顯示excel程式 
                excelApp.DisplayAlerts = false;        //设置是否显示警告窗体 
                book = excelApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
                Worksheet sheet = (Worksheet)book.Sheets[1]; 
                sheet.Activate();
            }
            else
            {
                object missing = System.Reflection.Missing.Value;
                book = this.ExcelAPP.Application.Workbooks.Open(strPath, missing, true, missing, missing, missing, missing, missing, missing, true, missing, missing, missing, missing, missing);

            }

            return book;

        }

        /// <summary>
        /// Excel 儲存資料使用的函數
        /// </summary>
        /// <param name="strPath"> 儲存檔案路徑 </param>
        /// <param name="Data"> 儲存資料 </param>
        /// <param name="poRow"> 啟始列(row)的位置 </param>
        /// <param name="poCol"> 啟始欄(column)的位置 </param>
        /// <param name="excelApp"> 目標儲存的excel app 物件 </param>
        /// <param name="sheet">目標儲存的excel sheet 物件 </param>
        /// <param name="book">目標儲存的excel book 物件 </param>
        private static void SaveFunction(string strPath, string[,] Data, int poRow, int poCol, Excel.Application excelApp, Worksheet sheet, Workbook book)
        {
            //// All into sheet
            int endRow = Data.GetLength(0) + poRow - 1;
            int endCol = Data.GetLength(1) + poCol - 1; 

            string StartPoString = GetPo(poCol);
            string EngPoString = GetPo(endCol);
            sheet.Activate();
            Range range = sheet.get_Range(StartPoString + poRow.ToString(),
                                          EngPoString + endRow.ToString());
            range.Value2 = Data;
            range.Value2 = range.Value2;


            book.SaveAs(strPath, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            sheet.SaveAs(strPath);


            book.Close();
            excelApp.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(book);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(sheet);
            GC.Collect();
        }

        /// <summary>
        /// 取得excel所有sheet名稱
        /// </summary>
        /// <returns></returns>
        private List<string> GetSheets()
        {
            int sheetCount = this.workBook.Sheets.Count;
            Worksheet ws = (Worksheet)this.workBook.Sheets[sheetCount];
            List<string> sheets = new List<string>();
            for (int i = 1; i < sheetCount + 1; i++)
            {
                Worksheet indSheet = (Worksheet)this.workBook.Sheets[i];
                sheets.Add(indSheet.Name);
            }
            return sheets;
        }

        /// <summary>
        /// 讀取excel資料使用的函數
        /// </summary>
        /// <param name="ws"> 讀取目標的excel sheet 物件 </param>
        /// <param name="RowStartPo"> 啟始列(row)的位置 </param>
        /// <param name="ColStartPo"> 啟始欄(column)的位置 </param> 
        /// <returns></returns>
        private static string[,] ReadGetData(Excel.Worksheet ws, int RowStartPo, int ColStartPo)
        {
            //取得总记录行数    (包括标题列)  
            int rowsint = ws.UsedRange.Cells.Rows.Count;            //得到列数    
            int columnsint = ws.UsedRange.Cells.Columns.Count;      //得到行数   
            //計算初始位置
            int[] startPo = new int[] { Convert.ToInt32(ColStartPo / 26), ColStartPo % 26 };
            string StartPo_ = GetPo(ColStartPo); 
            string StartPo = StartPo_ + RowStartPo.ToString();
            //計算結束位置
            columnsint = columnsint + ColStartPo - 1; 
            string EndPo_ = GetPo(columnsint);
            string EndPo = EndPo_ + (rowsint + RowStartPo - 1).ToString();
            //取的全部資料並儲存於arry1
            Excel.Range rng1 = ws.Cells.get_Range(StartPo, EndPo);
            object[,] arry1 = (object[,])rng1.Value2;
            int newRowNumber = arry1.GetLength(0);
            int newColNumber = arry1.GetLength(1);
            string[,] newData = new string[newRowNumber, newColNumber];

            for (int i = 1; i <= newRowNumber; i++)
            {
                for (int j = 1; j <= newColNumber; j++)
                {
                    try
                    {
                        newData[i - 1, j - 1] = arry1[i, j].ToString();
                    }
                    catch (Exception)
                    {
                        newData[i - 1, j - 1] = "";
                    }
                }
            }

            return newData;
        }

        /// <summary>
        /// 取得儲存Excel欄位的英文位置代號
        /// </summary>
        /// <param name="Num"></param>
        /// <returns></returns>
        private static string GetPo(int Num)
        {
            string po = "";
            string[] Engpo = new string[] { "", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z" };
            List<int> res = new List<int>();
            TransExcelPo(Num, ref res);
            res.Reverse();
            foreach (int ss in res)
            {
                po += Engpo[ss];
            }
            return po;
        }

        /// <summary>
        /// 此函數為遞迴函數，配合"GetPo"函數使用，將欲儲存的欄位數字轉換至26進位英文位置代號
        /// </summary>
        /// <param name="Num"> 欄位數字 </param>
        /// <param name="res"> 轉換後的代號位置數至 </param>
        private static void TransExcelPo(int Num, ref List<int> res)
        {
            if (Num == 0) return;

            if (Num % 26 != 0)
            {
                int rr = Num % 26;
                res.Add(rr);
                Num = (Num - rr) / 26 > 0 ? (Num - rr) / 26 : Num - rr;
                TransExcelPo(Num, ref res);
            }
            else
            {
                res.Add(Num);
            }
        }


         


        /////////////////// Will do

        //private void ThreadTest__GetAllSheetsData()
        //{
        //    List<Thread> THREADS = new List<Thread>();
        //    foreach (string ss in sheets)
        //    {
        //        Thread oThreadA = new Thread(new ParameterizedThreadStart());
        //        oThreadA.Name = ss;
        //        THREADS.Add(oThreadA);
        //    }

        //    for (int i = 0; i < sheets.Count(); i++)
        //    {
        //        var th = THREADS[i];
        //        th.Start(sheets[i]);
        //    }
        //}

        //private void ThreadTest__GetAllSheetsData_Process(object SheetName)
        //{
        //    DicSheetsData[SheetName.ToString()] = ReadBySheetName(1, 1, SheetName.ToString());
        //}
    }



    /////////////////////////////////////////////////////////////////////////////////////
    /////////////////////////////////////////////////////////////////////////////////////
    /////////////////////////////////////////////////////////////////////////////////////
    /////////////////////////////////////////////////////////////////////////////////////
    /////////////////////////////////////////////////////////////////////////////////////
    /////////////////////////////////////////////////////////////////////////////////////
    /////////////////////////////////////////////////////////////////////////////////////
    /////////////////////////////////////////////////////////////////////////////////////
    /////////////////////////////////////////////////////////////////////////////////////
    /////////////////////////////////////////////////////////////////////////////////////
    /////////////////////////////////////////////////////////////////////////////////////
    /////////////////////////////////////////////////////////////////////////////////////
    /////////////////////////////////////////////////////////////////////////////////////
    /////////////////////////////////////////////////////////////////////////////////////
    /////////////////////////////////////////////////////////////////////////////////////
    /////////////////////////////////////////////////////////////////////////////////////
    /////////////////////////////////////////////////////////////////////////////////////













    public class ExcelSaveAndRead
    {
        public static void SaveCreat(string strPath, string sheetName, int poRow, int poCol, string[,] Data)
        {
            bool fileExist = File.Exists(strPath);
            //// 若excel 不存在,創建
            if (!fileExist)
            {
                FileInfo fi = new FileInfo(strPath);
                Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
                excelApp.Visible = false;              //不顯示excel程式 
                excelApp.DisplayAlerts = false;        //设置是否显示警告窗体 
                Workbook book = excelApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
                Worksheet sheet = (Worksheet)book.Sheets[1];
                sheet.Name = sheetName;
                sheet.Activate();


                //// All into sheet
                int endRow = Data.GetLength(0) + poRow - 1;
                int endCol = Data.GetLength(1) + poCol - 1;
                string[] excelCol = new string[] { "", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z" };

                int endp2 = endCol % 26 == 0 ? 26 : endCol % 26;
                int endp1 = endp2 == 26 ? endCol / 26 - 1 : endCol / 26;
                int stp2 = poCol % 26 == 0 ? 26 : poCol % 26;
                int stp1 = stp2 == 26 ? poCol / 26 - 1 : poCol / 26;

                Range range = sheet.get_Range(excelCol[stp1] + excelCol[stp2] + poRow.ToString(), excelCol[endp1] + excelCol[endp2] + endRow.ToString());
                range.Value2 = Data;
                range.Value2 = range.Value2;



                book.SaveAs(strPath, Microsoft.Office.Interop.Excel.XlFileFormat.xlExcel8, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                book.Close(false, Type.Missing, Type.Missing);
                excelApp.Workbooks.Close();
                excelApp.Quit();
                //刪除 Windows工作管理員中的Excel.exe process，  
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(book);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(sheet);
                GC.Collect();
            }
            else
            {
                //// excel 檔案已存在
                //// 创建Application
                Excel.Application excelApp = new Excel.Application();
                excelApp.DisplayAlerts = false;
                excelApp.Visible = false;
                excelApp.ScreenUpdating = false;

                Excel.Workbook book = excelApp.Workbooks.Open(strPath);
                int sheetCount = book.Sheets.Count;
                Worksheet ws = (Worksheet)book.Sheets[sheetCount];
                Worksheet sheet = null;
                //// 檢查sheets 是否已存在,
                bool exist = false;
                for (int i = 1; i < sheetCount + 1; i++)
                {
                    Worksheet indSheet = (Worksheet)book.Sheets[i];
                    if (sheetName == indSheet.Name)
                    {
                        exist = true;
                        break;
                    }
                }

                //// 若sheet 已存在,刪除重新建立,否則建立sheet
                if (exist)
                {
                    ws = (Worksheet)book.Sheets[sheetCount];
                    sheet = (Worksheet)book.Worksheets.Add(Type.Missing, ws, Type.Missing, Type.Missing);//建立一個新分頁 
                    Worksheet sheet2 = book.Sheets[sheetName];
                    sheet2.Delete();
                    sheet.Name = sheetName;
                }
                else
                {
                    sheet = (Worksheet)book.Worksheets.Add(Type.Missing, ws, Type.Missing, Type.Missing);//建立一個新分頁 
                    sheet.Name = sheetName;
                }


                //// All into sheet
                int endRow = Data.GetLength(0) + poRow - 1;
                int endCol = Data.GetLength(1) + poCol - 1;
                string[] excelCol = new string[] { "", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z" };

                int endp2 = endCol % 26 == 0 ? 26 : endCol % 26;
                int endp1 = endp2 == 26 ? endCol / 26 - 1 : endCol / 26;
                int stp2 = poCol % 26 == 0 ? 26 : poCol % 26;
                int stp1 = stp2 == 26 ? poCol / 26 - 1 : poCol / 26;

                sheet.Activate();
                Range range = sheet.get_Range(excelCol[stp1] + excelCol[stp2] + poRow.ToString(), excelCol[endp1] + excelCol[endp2] + endRow.ToString());
                range.Value2 = Data;
                range.Value2 = range.Value2;


                book.SaveAs(strPath, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                sheet.SaveAs(strPath);

                book.Close();
                excelApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(book);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(sheet);
                GC.Collect();
            }
        }


        public static void Save(string strPath, int sheetNumber, int poRow, int poCol, string[,] Data)
        {
            //创建Application
            Excel.Application excelApp = new Excel.Application();
            //设置是否显示警告窗体
            excelApp.DisplayAlerts = false;
            //设置是否显示Excel
            excelApp.Visible = false;
            //禁止刷新屏幕
            excelApp.ScreenUpdating = false;
            // 加入新的活頁簿 

            //// All into sheet
            int endRow = Data.GetLength(0) + poRow - 1;
            int endCol = Data.GetLength(1) + poCol - 1;
            string[] excelCol = new string[] { "", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z" };

            int endp2 = endCol % 26 == 0 ? 26 : endCol % 26;
            int endp1 = endp2 == 26 ? endCol / 26 - 1 : endCol / 26;
            int stp2 = poCol % 26 == 0 ? 26 : poCol % 26;
            int stp1 = stp2 == 26 ? poCol / 26 - 1 : poCol / 26;

            Excel.Workbook book = excelApp.Workbooks.Open(strPath);
            Excel.Worksheet sheet = new Excel.Worksheet();
            sheet = book.Sheets[sheetNumber];
            sheet.Activate();
            Range range = sheet.get_Range(excelCol[stp1] + excelCol[stp2] + poRow.ToString(), excelCol[endp1] + excelCol[endp2] + endRow.ToString());
            range.Value2 = Data;
            range.Value2 = range.Value2;



            book.SaveAs(strPath, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            sheet.SaveAs(strPath);

            book.Close();
            excelApp.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(book);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(sheet);
            GC.Collect();

        }

        public static void Save(string strPath, string sheetNumber, int poRow, int poCol, string[,] Data)
        {

            //创建Application
            Excel.Application excelApp = new Excel.Application();
            //设置是否显示警告窗体
            excelApp.DisplayAlerts = false;
            //设置是否显示Excel
            excelApp.Visible = false;
            //禁止刷新屏幕
            excelApp.ScreenUpdating = false;


            //// All into sheet
            int endRow = Data.GetLength(0) + poRow - 1;
            int endCol = Data.GetLength(1) + poCol - 1;
            string[] excelCol = new string[] { "", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z" };

            int endp2 = endCol % 26 == 0 ? 26 : endCol % 26;
            int endp1 = endp2 == 26 ? endCol / 26 - 1 : endCol / 26;
            int stp2 = poCol % 26 == 0 ? 26 : poCol % 26;
            int stp1 = stp2 == 26 ? poCol / 26 - 1 : poCol / 26;

            Excel.Workbook book = excelApp.Workbooks.Open(strPath);
            Excel.Worksheet sheet = new Excel.Worksheet();
            sheet = book.Sheets[sheetNumber];
            sheet.Activate();
            Range range = sheet.get_Range(excelCol[stp1] + excelCol[stp2] + poRow.ToString(), excelCol[endp1] + excelCol[endp2] + endRow.ToString());
            range.Value2 = Data;
            range.Value2 = range.Value2;

            book.Close();
            excelApp.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(book);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(sheet);
            GC.Collect();
        }





        public static string[,] Read(string fullPath, int RowStartPo, int ColStartPo, int SheetNum)
        {
            string[] Engpo = new string[] { "", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z" };


            object missing = System.Reflection.Missing.Value;
            Excel.Application excel = new Excel.Application();//lauch excel application 
            excel.Visible = false; excel.UserControl = true;
            excel.DisplayAlerts = false;
            // 以只读的形式打开EXCEL文件  
            Excel.Workbook wb = excel.Application.Workbooks.Open(fullPath, missing, true, missing, missing, missing, missing, missing, missing, true, missing, missing, missing, missing, missing);
            //取得第 SheetNum 个工作薄  


            Excel.Worksheet ws = (Excel.Worksheet)wb.Worksheets.get_Item(SheetNum);
            string[,] newData = ReadGetData(ws, RowStartPo, ColStartPo, Engpo);


            excel.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(wb);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(ws);
            GC.Collect();
            return newData;
        }

        public static string[,] ReadBySheetName(string fullPath, int RowStartPo, int ColStartPo, string SheetName)
        {
            string[] Engpo = new string[] { "", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z" };
            object missing = System.Reflection.Missing.Value;
            Excel.Application excel = new Excel.Application();//lauch excel application 
            excel.Visible = false; excel.UserControl = true;
            excel.DisplayAlerts = false;
            // 以只读的形式打开EXCEL文件  
            Excel.Workbook wb = excel.Application.Workbooks.Open(fullPath, missing, true, missing, missing, missing, missing, missing, missing, true, missing, missing, missing, missing, missing);
            //取得第 SheetNum 个工作薄  

            List<string> sheets = GetSheets(wb);
            var SheetNum = sheets.IndexOf(SheetName) + 1;

            Excel.Worksheet ws = (Excel.Worksheet)wb.Worksheets.get_Item(SheetNum);
            string[,] newData = ReadGetData(ws, RowStartPo, ColStartPo, Engpo);


            excel.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(wb);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(ws);
            GC.Collect();
            return newData;
        }

        private static List<string> GetSheets(Excel.Workbook book)
        {


            int sheetCount = book.Sheets.Count;
            Worksheet ws = (Worksheet)book.Sheets[sheetCount];
            Worksheet sheet = null;
            //// 檢查sheets 是否已存在,
            bool exist = false;
            List<string> sheets = new List<string>();
            for (int i = 1; i < sheetCount + 1; i++)
            {
                Worksheet indSheet = (Worksheet)book.Sheets[i];
                sheets.Add(indSheet.Name);
            }

            return sheets;
        }



        public static List<string> GetSheets(string strPath)
        {

            Excel.Application excelApp = new Excel.Application();
            excelApp.DisplayAlerts = false;
            excelApp.Visible = false;
            excelApp.ScreenUpdating = false;

            Excel.Workbook book = excelApp.Workbooks.Open(strPath);
            int sheetCount = book.Sheets.Count;
            Worksheet ws = (Worksheet)book.Sheets[sheetCount];
            Worksheet sheet = null;
            //// 檢查sheets 是否已存在,
            bool exist = false;
            List<string> sheets = new List<string>();
            for (int i = 1; i < sheetCount + 1; i++)
            {
                Worksheet indSheet = (Worksheet)book.Sheets[i];
                sheets.Add(indSheet.Name);
            }

            book.Close();
            excelApp.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(book);
            GC.Collect();
            return sheets;
        }





        public static List<string[,]> Read(string fullPath, int RowStartPo, int ColStartPo, string SheetNumAll)
        {
            string[] Engpo = new string[] { "", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z" };


            object missing = System.Reflection.Missing.Value;
            Excel.Application excel = new Excel.Application();//lauch excel application 
            excel.Visible = false;
            excel.UserControl = true;
            excel.DisplayAlerts = false;
            // 以只读的形式打开EXCEL文件  
            Excel.Workbook wb = excel.Application.Workbooks.Open(fullPath, missing, true, missing, missing, missing, missing, missing, missing, true, missing, missing, missing, missing, missing);
            //取得第 SheetNum 个工作薄  



            string[] sheets = SheetNumAll.IndexOf(",") != -1 ? SheetNumAll.Split(',') :
                              (SheetNumAll.IndexOf("-") != -1 ? SheetNumAll.Split('-') : (new string[1] { SheetNumAll }));

            List<string[,]> AllData = new List<string[,]>();
            if (SheetNumAll.IndexOf("-") != -1)
            {
                for (int i = Int32.Parse(sheets[0]); i <= Int32.Parse(sheets[1]); i++)
                {
                    Excel.Worksheet ws = (Excel.Worksheet)wb.Worksheets.get_Item(i);
                    string[,] newData = ReadGetData(ws, RowStartPo, ColStartPo, Engpo);
                    AllData.Add(newData);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ws);
                }
            }
            else
            {
                foreach (string ss in sheets)
                {
                    Excel.Worksheet ws = (Excel.Worksheet)wb.Worksheets.get_Item(Int32.Parse(ss));
                    string[,] newData = ReadGetData(ws, RowStartPo, ColStartPo, Engpo);
                    AllData.Add(newData);
                }
            }

            excel.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(wb);
            GC.Collect();
            return AllData;
        }









        private static string[,] ReadGetData(Excel.Worksheet ws, int RowStartPo, int ColStartPo, string[] Engpo)
        {
            //取得总记录行数    (包括标题列)  
            int rowsint = ws.UsedRange.Cells.Rows.Count;            //得到列数    
            int columnsint = ws.UsedRange.Cells.Columns.Count;      //得到行数   
            //計算初始位置
            int[] startPo = new int[] { Convert.ToInt32(ColStartPo / 26), ColStartPo % 26 };
            startPo[0] = startPo[1] == 0 ? startPo[0] - 1 : startPo[0];
            startPo[1] = startPo[1] == 0 ? 26 : startPo[1];
            string StartPo = Engpo[startPo[0]] + Engpo[startPo[1]] + RowStartPo.ToString();
            //計算結束位置
            columnsint = columnsint + ColStartPo - 1;
            int[] endPo = new int[] { Convert.ToInt32(columnsint / 26), columnsint % 26 };
            endPo[0] = endPo[1] == 0 ? endPo[0] - 1 : endPo[0];
            endPo[1] = endPo[1] == 0 ? 26 : endPo[1];
            string EndPo = Engpo[endPo[0]] + Engpo[endPo[1]] + (rowsint + RowStartPo - 1).ToString();
            //取的全部資料並儲存於arry1
            Excel.Range rng1 = ws.Cells.get_Range(StartPo, EndPo);
            object[,] arry1 = (object[,])rng1.Value2;
            int newRowNumber = arry1.GetLength(0);
            int newColNumber = arry1.GetLength(1);
            string[,] newData = new string[newRowNumber, newColNumber];

            for (int i = 1; i <= newRowNumber; i++)
            {
                for (int j = 1; j <= newColNumber; j++)
                {
                    try
                    {
                        newData[i - 1, j - 1] = arry1[i, j].ToString();
                    }
                    catch (Exception)
                    {
                        newData[i - 1, j - 1] = "";
                    }
                }
            }

            return newData;
        }










    }
}
