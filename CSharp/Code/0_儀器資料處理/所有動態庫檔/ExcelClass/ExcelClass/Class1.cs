using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;

namespace ExcelClass
{
    public class ExcelSaveAndRead
    {
        public static void SaveCreat(string strPath, string sheetName, int poRow, int poCol, string[,] Data)
        { 
            bool fileExist = File.Exists(strPath);
            //// 若excel 不存在,創建
            if ( !fileExist)
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
                for (int i = 1; i < sheetCount+1; i++)
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









        private static string[,] ReadGetData(Excel.Worksheet ws, int RowStartPo , int ColStartPo, string[] Engpo)
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
