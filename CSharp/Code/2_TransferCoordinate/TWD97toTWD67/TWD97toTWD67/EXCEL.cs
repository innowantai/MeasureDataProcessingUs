using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace TWD97toTWD67
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
        /// 讀取指定sheet名稱的資料
        /// </summary>
        /// <param name="RowStartPo"> 指定列(row)的開始位置 </param>
        /// <param name="ColStartPo"> 指定欄(column)的開始位置 </param>
        /// <param name="SheetName"> 指定excel分頁的名稱 </param>
        /// <returns></returns>
        public string[,] GetDataBySheetNumber(int RowStartPo, int ColStartPo, int SheetNnumber)
        {
            Excel.Worksheet ws = (Excel.Worksheet)this.workBook.Worksheets.get_Item(SheetNnumber);
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
        /// 儲存資料至此開啟excel檔案
        /// </summary>
        /// <param name="sheetName"> 指定儲存之分頁名稱 </param>
        /// <param name="poRow"> 啟始列(row)的位置 </param>
        /// <param name="poCol"> 啟始欄(column)的位置 </param>
        /// <param name="Data"> 儲存資料 </param>
        public void Save_ChangeFormat(string sheetName, int poRow, int poCol, string[,] Data)
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

                SaveFunction_ChangeFormat(strPath, Data, poRow, poCol, excelApp, sheet, book);

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

                SaveFunction_ChangeFormat(strPath, Data, poRow, poCol, excelApp, sheet, book);

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
        private void SaveFunction(string strPath, string[,] Data, int poRow, int poCol, Excel.Application excelApp, Worksheet sheet, Workbook book)
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
        /// Excel 儲存資料使用的函數
        /// </summary>
        /// <param name="strPath"> 儲存檔案路徑 </param>
        /// <param name="Data"> 儲存資料 </param>
        /// <param name="poRow"> 啟始列(row)的位置 </param>
        /// <param name="poCol"> 啟始欄(column)的位置 </param>
        /// <param name="excelApp"> 目標儲存的excel app 物件 </param>
        /// <param name="sheet">目標儲存的excel sheet 物件 </param>
        /// <param name="book">目標儲存的excel book 物件 </param>
        private void SaveFunction_ChangeFormat(string strPath, string[,] Data, int poRow, int poCol, Excel.Application excelApp, Worksheet sheet, Workbook book)
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


            Style_FontSize(sheet, 1, 1, 1, 12, 18);
            Style_FontBord(sheet, 1, 1, 1, 12);
            Style_Merge(sheet, 1, 1, 1, 12);
            Style_Alignment(sheet, 1, 1, 1, 12);

            Style_FontSize(sheet, 2, 1, 1, 1, 14);
            Style_FontBord(sheet, 2, 1, 1, 1);

            Style_FontSize(sheet, 3, 1, 1, 1, 14);
            Style_FontBord(sheet, 3, 1, 1, 1);


            int Row_Len = Data.GetLength(0);
            Style_BorderWeight(sheet, 6, 1, Row_Len - 12, 9, 2);


            int btn_Po = Data.GetLength(0) - 4; 
            Style_BorderWeightWithDirection(sheet, btn_Po, 1, 5, 6, 4, "LL");
            Style_BorderWeightWithDirection(sheet, btn_Po, 1, 5, 6, 4, "TT");
            Style_BorderWeightWithDirection(sheet, btn_Po, 1, 5, 6, 4, "RR");
            Style_BorderWeightWithDirection(sheet, btn_Po, 1, 5, 6, 4, "BB");
            Style_FontSize(sheet, btn_Po, 1, 4, 6, 14);
            Style_FontBord(sheet, btn_Po, 1, 1, 1);
            Style_FontBord(sheet, btn_Po + 3, 1, 1, 10); 
            Style_Merge(sheet, btn_Po, 1, 1, 2);
            Style_Merge(sheet, btn_Po + 1, 1, 2, 4);


            book.SaveAs(strPath, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            sheet.SaveAs(strPath);


            book.Close();
            excelApp.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(book);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(sheet);
            GC.Collect();
        }

        private void Style_BorderWeight(Worksheet sheet, int poRow, int poCol, int Len_row, int Len_Col, int Weight)
        {
            //// All into sheet
            int endRow = Len_row + poRow - 1;
            int endCol = Len_Col + poCol - 1;

            string StartPoString = GetPo(poCol);
            string EngPoString = GetPo(endCol);

            Range ra = sheet.get_Range(StartPoString + poRow.ToString(),
                                       EngPoString + endRow.ToString());

            ra.Borders.Weight = Weight;
            ra.Columns.AutoFit();
            //ra.Borders.Weight = Excel.XlBorderWeight.xlThin; 
        }


        private void Style_BorderWeightWithDirection(Worksheet sheet, int poRow, int poCol, int Len_row, int Len_Col, int Weight, string Dir)
        {
            //// All into sheet
            int endRow = Len_row + poRow - 1;
            int endCol = Len_Col + poCol - 1;

            string StartPoString = GetPo(poCol);
            string EngPoString = GetPo(endCol);

            Range ra = sheet.get_Range(StartPoString + poRow.ToString(),
                                       EngPoString + endRow.ToString());
            if (Dir == "LL")
            {
                ra.Borders[XlBordersIndex.xlEdgeLeft].Weight = Weight;
            }
            else if (Dir == "RR")
            {
                ra.Borders[XlBordersIndex.xlEdgeRight].Weight = Weight;
            }
            else if (Dir == "TT")
            {
                ra.Borders[XlBordersIndex.xlEdgeTop].Weight = Weight;
            }
            else
            {
                ra.Borders[XlBordersIndex.xlEdgeBottom].Weight = Weight;
            } 
            ra.Columns.AutoFit();
            //ra.Borders.Weight = Excel.XlBorderWeight.xlThin; 
        }




        private void Style_Alignment(Worksheet sheet, int poRow, int poCol, int Len_row, int Len_Col)
        {
            //// All into sheet
            int endRow = Len_row + poRow - 1;
            int endCol = Len_Col + poCol - 1;

            string StartPoString = GetPo(poCol);
            string EngPoString = GetPo(endCol);

            Range ra = sheet.get_Range(StartPoString + poRow.ToString(),
                                       EngPoString + endRow.ToString());
            ra.HorizontalAlignment = Excel.XlVAlign.xlVAlignCenter;
        }


        private void Style_FontSize(Worksheet sheet, int poRow, int poCol, int Len_row, int Len_Col, int fontSize)
        {
            //// All into sheet
            int endRow = Len_row + poRow - 1;
            int endCol = Len_Col + poCol - 1;

            string StartPoString = GetPo(poCol);
            string EngPoString = GetPo(endCol);

            Range ra = sheet.get_Range(StartPoString + poRow.ToString(),
                                       EngPoString + endRow.ToString());
            ra.Font.Size = fontSize;
        }

        private void Style_FontBord(Worksheet sheet, int poRow, int poCol, int Len_row, int Len_Col)
        {
            //// All into sheet
            int endRow = Len_row + poRow - 1;
            int endCol = Len_Col + poCol - 1;

            string StartPoString = GetPo(poCol);
            string EngPoString = GetPo(endCol);

            Range ra = sheet.get_Range(StartPoString + poRow.ToString(),
                                       EngPoString + endRow.ToString());
            ra.Font.Bold = 4;
            ra.Columns.AutoFit();
        }


        private void Style_Merge(Worksheet sheet, int poRow, int poCol, int Len_row, int Len_Col)
        {
            //// All into sheet
            int endRow = Len_row + poRow - 1;
            int endCol = Len_Col + poCol - 1;

            string StartPoString = GetPo(poCol);
            string EngPoString = GetPo(endCol);

            Range ra = sheet.get_Range(StartPoString + poRow.ToString(),
                                       EngPoString + endRow.ToString());
            ra.Merge();
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
        private string[,] ReadGetData(Excel.Worksheet ws, int RowStartPo, int ColStartPo)
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

            int Last_I = 0;
            int Last_J = 0;
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

                    if (newData[i - 1, j - 1] != "")
                    {
                        Last_I = i - 1 > Last_I ? i - 1 : Last_I;
                        Last_J = j - 1 > Last_J ? j - 1 : Last_J;
                    }
                }
            }

            string[,] resData = new string[Last_I + 1, Last_J + 1];
            for (int i = 0; i <= Last_I; i++)
            {
                for (int j = 0; j <= Last_J; j++)
                {
                    resData[i, j] = newData[i, j];
                }
            }






            return resData;
        }

        /// <summary>
        /// 取得儲存Excel欄位的英文位置代號
        /// </summary>
        /// <param name="Num"></param>
        /// <returns></returns>
        private string GetPo(int Num)
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
        private void TransExcelPo(int Num, ref List<int> res)
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


}
