using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Collections;
using System.IO;

/*
 * 201806150835
 1.有兩種資料格式,自動判斷並轉換格式
 2.資料可能錯誤類型
     (1).資料順序Rb,Rf,Rf,Rb 排序有誤  --> 可能是缺資料
     (2).Start-Lint 與 End-Line 數不同 --> 紀錄可能斷掉
     (3).相同點位往返測量之高程差過大  --> 大於所設定之標準 
 */

namespace OBDAT
{
    public class OBDAT
    {
        public static string oriPath = System.Environment.CurrentDirectory;
        public static string Start1, Start2;
        public static string outPutTxt = "";
        public static double diffStandard; 



        public static string OBMain_sub(string oriPath, string savePath, string fileName, string Standard)
        {
            outPutTxt = "";
            string dataPath = Path.Combine(savePath, fileName);
            try
            {
                diffStandard = Convert.ToDouble(Standard.Trim());
            }
            catch (Exception)
            {

                return "誤差值設置有誤 \r\n";
            }


            //// Loading Data and Data position save to the variable of Data and position
            ArrayList Data = new ArrayList();
            ArrayList position = new ArrayList();
            Loading(ref position, ref Data, dataPath);

            bool Error1 = ErrorCheck1_sub(Data, position);
            if (Error1) { return outPutTxt; }


            //// Judge data format type and transfer to processable type.
            string index = (System.String)Data[0];
            int caseNum = index.Contains("|") ? 0 : 1;
            Data = FormateCase2toCase1(Data, caseNum);

            string OutSaveName = null;
            ArrayList sortData = OUT_sub(Data, savePath, ref OutSaveName, caseNum);
            OutSaveName = OutSaveName.Replace(".out", ".xls");
            string[,] excelRes = ExcelProcess(sortData);

            File.Delete(Path.Combine(savePath, OutSaveName));
            File.Copy(Path.Combine(oriPath, "Ori_excel.xls"), Path.Combine(savePath, OutSaveName));
            ExcelClass.ExcelSaveAndRead.Save(strPath: Path.Combine(savePath, OutSaveName), sheetNumber: 9, poRow: 2, poCol: 1, Data: excelRes);


            outPutTxt += "    處理完成 \r\n";
            return outPutTxt;

        }







        /// <summary>
        /// Arrange to .OUT Form
        /// </summary>
        /// <param name="Data"></param>
        /// <param name="savePath"></param>
        /// <param name="OutSaveName"></param>
        /// <returns></returns>
        public static ArrayList OUT_sub(ArrayList Data, string savePath, ref string OutSaveName, int caseNum)
        {
            //OutSaveName = Convert.ToString(Data[0]);
            //if (caseNum == 0)
            //{
            //}
            //else
            //{
            //    OutSaveName = OutSaveName.Trim();
            //    OutSaveName = OutSaveName.Substring(2, OutSaveName.Length - 2).Trim();
            //    OutSaveName = OutSaveName.Replace(".dat", ".out");
            //    OutSaveName = OutSaveName.Replace(".DAT", ".out");
            //}
            OutSaveName = Convert.ToString(Data[0]);
            OutSaveName = OutSaveName.Substring(21, 15).Trim();
            OutSaveName = OutSaveName.Replace(".dat", ".out");
            OutSaveName = OutSaveName.Replace(".DAT", ".out");

            //// Sort Data  
            ArrayList RbRfData = new ArrayList();
            ArrayList allData = new ArrayList();
            string Sh, Db, End;
            int i = 0;
            foreach (string ff in Data)
            {
                if (ff.Contains("Start-Line"))
                {
                    Start1 = Convert.ToString(Data[i]).Replace("BFFB", "  BF");
                    Start2 = Convert.ToString(Data[i + 1]);
                    RbRfData.Clear();
                }
                else if (ff.Contains("|Rb") || ff.Contains("|Rf"))
                {
                    RbRfData.Add(ff);
                }
                else if (ff.Contains("|Sh"))
                {
                    Sh = Convert.ToString(Data[i]);
                    Db = Convert.ToString(Data[i + 1]);
                    End = Convert.ToString(Data[i + 2]);

                    string GoName = Start2.Substring(22, 9);
                    string BackName = Db.Substring(22, 9);

                    //// Go
                    //// Save Start-Line
                    allData.Add(Start1);
                    allData.Add(Start2);
                    //// Save Rb Rf data
                    for (int ii = 0; ii < RbRfData.Count / 4; ii++)
                    {
                        allData.Add(RbRfData[ii * 4]);
                        allData.Add(RbRfData[ii * 4 + 1]);
                    }
                    //// Save Sh,Db,End-Line
                    allData.Add(Sh);
                    allData.Add(Db);
                    allData.Add(End);

                    //// Back
                    //// Save Start-Line
                    allData.Add(Start1.Replace(GoName, BackName));
                    allData.Add(Start2.Replace(GoName, BackName));
                    //// Save Rb, Rf Data
                    for (int ii = 0; ii < RbRfData.Count / 4; ii++)
                    {
                        String index1 = Convert.ToString(RbRfData[RbRfData.Count - ii * 4 - 1 - 1]);
                        String index2 = Convert.ToString(RbRfData[RbRfData.Count - ii * 4 - 1]);
                        index1 = index1.Replace("Rf", "Rb");
                        index2 = index2.Replace("Rb", "Rf");
                        allData.Add(index1);
                        allData.Add(index2);
                    }
                    //// Save Sh,Db,End-line
                    allData.Add(Sh.Replace(BackName, GoName));
                    allData.Add(Db.Replace(BackName, GoName));
                    allData.Add(End);
                }
                i++;
            }
            StreamWriter sw = new StreamWriter(Path.Combine(savePath, OutSaveName));
            foreach (string ff in allData)
            {
                sw.WriteLine(ff);
            }
            sw.WriteLine("-9999");
            sw.Close();

            return allData;

        }

        /// <summary>
        /// Sort to Excel Form
        /// </summary>
        /// <param name="sortData"></param>
        /// <returns></returns>
        public static string[,] ExcelProcess(ArrayList sortData)
        {
            string[,] excelData = new string[sortData.Count, 5];
            string StationName;
            int kk = 0;
            int CASE = 0;
            foreach (string ff in sortData)
            {

                if (ff.Contains("Start"))
                {
                    excelData[kk, 4] = CASE % 2 == 0 ? "往程觀測" : "返程觀測";
                    excelData[kk, 2] = "0.00000";
                }
                else if (ff.Contains("End"))
                {
                    kk++;
                    excelData[kk, 0] = "END";
                    excelData[kk - 1, 1] = "0.00000";
                    excelData[kk - 1, 3] = "0.00000";
                    CASE++;
                    kk++;
                }
                else if (ff.Contains("Rb"))
                {
                    StationName = ff.Substring(22, 13).Trim();
                    string BHD = ff.Substring(77, 13).Trim();
                    string RB = ff.Substring(53, 14).Trim();
                    excelData[kk, 0] = StationName;
                    excelData[kk, 1] = BHD;
                    excelData[kk, 3] = RB;
                }
                else if (ff.Contains("Rf"))
                {
                    StationName = ff.Substring(22, 13).Trim();
                    string FHD = ff.Substring(77, 13).Trim();
                    string RF = ff.Substring(53, 14).Trim();
                    excelData[kk + 1, 0] = StationName;
                    excelData[kk + 1, 2] = FHD;
                    excelData[kk + 1, 4] = RF;
                    kk++;
                }
            }
            return excelData;
        }

        /// <summary>
        /// 將第二種格式轉換排列成第一種格式輸出
        /// </summary>
        /// <param name="data"></param>
        /// <param name="caseNum"></param>
        /// <returns></returns>
        private static ArrayList FormateCase2toCase1(ArrayList data, int caseNum)
        {
            ArrayList newData = new ArrayList();
            if (caseNum == 1)
            {
                int kk = 0;
                string[,] indexData = new string[data.Count, 10];
                string[] block = new string[] { "", " ", "  ", "   ", "    ", "     ", "      ", "       ", "        ", "         ", "          ", "           ", "            ", "             ", "              ", "               ", "                " };
                int[] mark = new int[data.Count];
                foreach (string ff in data)
                {
                    string index = ff.Trim();
                    string index1, index2, index3, index4, index5;
                    mark[kk] = ff.Contains("Start") ? 0 : (ff.Contains("End") ? 1 : (ff.Contains("Rb") | ff.Contains("Rf") ? 2 : (ff.Contains("Sh") ? 3 : (ff.Contains("Db") ? 4 : (ff.Contains("Z") ? 5 : 6)))));

                    if (mark[kk] == 0)
                    {
                        indexData[kk, 0] = "To  ";
                        indexData[kk, 1] = "Start-Line";
                        indexData[kk, 2] = "       BFFB     1|                      |                      |                      | ";
                        indexData[kk, 3] = "";
                    }
                    else if (mark[kk] == 1)
                    {
                        indexData[kk, 0] = "To  End-Line                  1|                      |                      |                      | ";
                    }
                    else if (mark[kk] == 2)
                    {
                        int po1 = index.IndexOf(" ");
                        int po2 = index.IndexOf(":");
                        int po3 = index.IndexOf(" Rb") == -1 ? index.IndexOf(" Rf") : index.IndexOf(" Rb");
                        int po4 = index.IndexOf("HD");
                        int po5 = index.IndexOf("sR");
                        string sign = index.IndexOf("Rb") == -1 ? "Rf" : "Rb";
                        index1 = index.Substring(po1, po2 - po1 - 4).Trim();
                        index2 = index.Substring(po2 - 2, po3 - po2).Trim();
                        index3 = index.Substring(po3 + 3, po4 - po3 - 3).Trim();
                        index4 = index.Substring(po4 + 3, po5 - po4 - 3).Trim();
                        indexData[kk, 0] = "KD1";
                        indexData[kk, 1] = block[9 - index1.Length] + index1;
                        indexData[kk, 2] = "      " + index2 + block[12 - index2.Length];
                        indexData[kk, 3] = "1|" + sign + block[15 - index3.Length] + index3 + " m";
                        indexData[kk, 4] = "   |HD" + block[15 - index4.Length] + index4 + " m   |                      |";
                    }
                    else if (mark[kk] == 3)
                    {
                        int po1 = index.IndexOf(" ");
                        int po2 = index.IndexOf("Sh");
                        int po3 = index.IndexOf("dz");
                        int po4 = index.IndexOf("Z");
                        index1 = index.Substring(po1, po2 - po1 - 4).Trim();
                        index2 = index.Substring(po2 + 2, po3 - po2 - 2).Trim();
                        index3 = index.Substring(po3 + 2, po4 - po3 - 2).Trim();
                        index4 = index.Substring(po4 + 2, index.Length - po4 - 2).Trim();
                        indexData[kk, 0] = "KD1";
                        indexData[kk, 1] = block[9 - index1.Length] + index1;
                        indexData[kk, 3] = "                  1|Sh" + block[15 - index2.Length] + index2 + " m";
                        indexData[kk, 4] = "   |dz" + block[15 - index3.Length] + index3 + " m   |Z"; ;
                        indexData[kk, 5] = block[16 - index4.Length] + index4 + " m   |";
                    }
                    else if (mark[kk] == 4)
                    {
                        int po1 = index.IndexOf(" ");
                        int po2 = index.IndexOf("Db");
                        int po3 = index.IndexOf("Df");
                        int po4 = index.IndexOf("Z");
                        string tmpIndex = index.Substring(po1, po2 - po1 - 4).Trim();
                        index1 = tmpIndex.Substring(0, tmpIndex.Length - 1).Trim();
                        index2 = tmpIndex.Substring(tmpIndex.IndexOf(" "), tmpIndex.Length - tmpIndex.IndexOf(" ")).Trim();
                        index3 = index.Substring(po2 + 2, po3 - po2 - 2).Trim();
                        index4 = index.Substring(po3 + 2, po4 - po3 - 2).Trim();
                        index5 = index.Substring(po4 + 2, index.Length - po4 - 2).Trim();
                        indexData[kk, 0] = "KD1";
                        indexData[kk, 1] = block[9 - index1.Length] + index1;
                        indexData[kk, 2] = "      " + index2 + block[12 - index2.Length];
                        indexData[kk, 3] = "1|Rb" + block[15 - index3.Length] + index3 + " m";
                        indexData[kk, 4] = "   |Df" + block[15 - index4.Length] + index4 + " m   |Z";
                        indexData[kk, 5] = block[16 - index5.Length] + index5 + " m   |";
                    }
                    else if (mark[kk] == 5)
                    {
                        int po1 = index.IndexOf(" ");
                        int po2 = index.IndexOf("Z");
                        string tmpIndex = index.Substring(po1, index.Length - po1).Trim();
                        index1 = tmpIndex.Substring(0, tmpIndex.IndexOf(" ")).Trim();
                        index2 = index.Substring(po2 + 1, index.Length - po2 - 1).Trim();
                        indexData[kk, 0] = "KD1";
                        indexData[kk, 1] = block[9 - index1.Length] + index1 + "                  1|                      |                      |Z";
                        indexData[kk, 5] = block[16 - index2.Length] + index2 + " m   |";
                    }
                    else if (mark[kk] == 6)
                    {
                        index1 = index.Substring(index.IndexOf(" "), index.Length - index.IndexOf(" ")).Trim();
                        indexData[kk, 0] = "To  ";
                        indexData[kk, 1] = index1 + block[27 - index1.Length] + "|                      |                      |                      | ";
                    }



                    string tmp = "For M5|Adr" + block[6 - (kk + 1).ToString().Length] + (kk + 1).ToString() + "|";
                    for (int pp = 0; pp < 6; pp++)
                    {
                        string tmpIndex = indexData[kk, pp];
                        tmp += tmpIndex;
                    }
                    newData.Add(tmp);
                    kk++;
                }
                return newData;
            }
            return data;
        }



        /// <summary>
        /// 檢查資料組數是否為Rb-Rf-Rf-Rb 排列
        /// </summary>
        /// <param name="Data"></param>
        /// <param name="position"></param>
        /// <returns></returns>
        public static bool ErrorCheck1_sub(ArrayList Data, ArrayList position)
        {
            ArrayList ErrorData = new ArrayList();
            ArrayList ErrorPo = new ArrayList();
            //// Initial Error Variable that is BOOL type to confirm Error state
            bool Error = false;
            //// Find All RB , RF and that position in ori file 
            int ii = 0;
            int[] kk = new int[4];
            foreach (string ff in Data)
            {
                //// Check Error1 : Rb Rf
                if (ff.IndexOf("Rb") != -1 || ff.IndexOf("Rf") != -1)
                {
                    string index = ff.Substring(ff.IndexOf("R"), 19);
                    ErrorData.Add(ff);
                    ErrorPo.Add(position[ii]);
                }
                ii++;

                //// Check Error2 : the series of go and back
                if (ff.IndexOf("Start-Line") != -1)
                {
                    kk[0] += 1;
                }
                else if (ff.IndexOf("Sh") != -1)
                {
                    kk[1] += 1;
                }
                else if (ff.IndexOf("Db") != -1)
                {
                    kk[2] += 1;
                }
                else if (ff.IndexOf("End-Line") != -1)
                {
                    kk[3] += 1;
                }
            }


            //// Error1 : Check the 4 serise-data iter is the Rb-Rf-Rf-Rb ? OK or NG(report NG-data position)
            int len = ErrorData.Count;
            for (int i = 0; i < len; i += 4)
            {
                int index1 = ErrorData[i].ToString().IndexOf("Rb");
                int index2 = ErrorData[i + 1].ToString().IndexOf("Rf");
                int index3 = ErrorData[i + 2].ToString().IndexOf("Rf");
                int index4 = ErrorData[i + 3].ToString().IndexOf("Rb");
                int check1 = index1 == -1 ? 1 : (index2 == -1 ? 2 : (index3 == -1 ? 3 : (index4 == -1 ? 4 : 0)));
                if (check1 != 0)
                {
                    outPutTxt += "測站組數有誤,可能錯誤位置為 : \r\n";
                    Error = true;
                    if (index1 == -1)
                    {
                        for (int jj = 0; jj < 8; jj++)
                        {
                            int Po = Convert.ToInt32(ErrorPo[i - 4 + jj]) + 1;
                            outPutTxt += "    第" + Po.ToString() + "列 : " + ErrorData[i - 4 + jj] + "\r\n";
                        }
                    }
                    else
                    {
                        for (int jj = 0; jj < 4; jj++)
                        {
                            int Po = Convert.ToInt32(ErrorPo[i + jj]) + 1;
                            outPutTxt += "    第" + Po.ToString() + "列 : " + ErrorData[i - 4 + jj] + "\r\n";
                        }
                    }
                    return Error;
                }
            }

            //// Error 2
            bool check2 = kk.Min() != kk.Max() ? true : false;
            if (check2)
            {
                Error = true;
                outPutTxt += kk[0] + "," + kk[1] + "," + kk[2] + "," + kk[3] + "往返組數有誤 \r\n";
                return Error;
            }


            //// Error3 Check height-diff 
            bool err3Check = false;
            if (!Error)
            {
                outPutTxt += "往返高程差大於標準(" + diffStandard.ToString() + ") \r\n";
                for (int i = 0; i < len; i += 4)
                {
                    string ind1 = ErrorData[i].ToString();
                    int po1 = ind1.IndexOf("Rb") == -1 ? ind1.IndexOf("Rf") : ind1.IndexOf("Rb");
                    int po2 = ind1.IndexOf("HD");
                    po1 = po1 + 2;
                    string ind2 = ErrorData[i].ToString().Substring(po1, po2 - po1).Trim();
                    double index1 = Convert.ToDouble(ErrorData[i].ToString().Substring(po1, po2 - po1).Replace("|", "").Replace("m", "").Trim());
                    double index2 = Convert.ToDouble(ErrorData[i + 1].ToString().Substring(po1, po2 - po1).Replace("|", "").Replace("m", "").Trim());
                    double index3 = Convert.ToDouble(ErrorData[i + 2].ToString().Substring(po1, po2 - po1).Replace("|", "").Replace("m", "").Trim());
                    double index4 = Convert.ToDouble(ErrorData[i + 3].ToString().Substring(po1, po2 - po1).Replace("|", "").Replace("m", "").Trim());
                    //double index1 = Convert.ToDouble(ErrorData[i].ToString().Substring(59, 8).Trim());
                    //double index2 = Convert.ToDouble(ErrorData[i + 1].ToString().Substring(59, 8).Trim());
                    //double index3 = Convert.ToDouble(ErrorData[i + 2].ToString().Substring(59, 8).Trim());
                    //double index4 = Convert.ToDouble(ErrorData[i + 3].ToString().Substring(59, 8).Trim());
                    double diff1 = Math.Abs(index2 - index1);
                    double diff2 = Math.Abs(index3 - index4);
                    double diff = Math.Abs(diff1 - diff2);
                    if (diff >= diffStandard)
                    {
                        for (int jj = 0; jj < 4; jj++)
                        {
                            int Po = Convert.ToInt32(ErrorPo[i + jj]) + 1;
                            outPutTxt += "    第" + Po.ToString() + "列 : " + ErrorData[i + jj] + "\r\n";
                        }
                        err3Check = true;
                    }

                    /* Console.WriteLine(index1);
                     Console.WriteLine(index2);
                     Console.WriteLine(index3);
                     Console.WriteLine(index4);
                     Console.WriteLine("---- " + diff1);
                     Console.WriteLine("---- " + diff2);
                     Console.WriteLine("---- " + diff);*/
                }

            }
            outPutTxt = err3Check ? outPutTxt : "";
            return Error;
        }

        /// <summary>
        /// 讀檔
        /// </summary>
        /// <param name="position"></param>
        /// <param name="Data"></param>
        public static void Loading(ref ArrayList position, ref ArrayList Data, string dataPath)
        {
            StreamReader sr = new StreamReader(@dataPath);
            int i = 0;
            while (!sr.EndOfStream)
            {
                string index = sr.ReadLine();
                if (!index.Contains("#") && !index.Contains("repeated"))
                {
                    Data.Add(index);
                    position.Add(i);
                }
                i++;
            }
            sr.Close();
        }

    }
}
