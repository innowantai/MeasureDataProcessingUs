using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Collections;

/*
 資料可能錯誤類型
 1.Car,Geo,Pro, 三種座標系統的資料總數有誤
 2.Car,Geo,Pro, 三種座標系統各自的測站數與資料數不符
 3.station to station 與 Measurement (Project......) 的資料總數不同 
 */

namespace Assembled
{
    public static class GPS_RES
    {

        public static string oriPath = System.Environment.CurrentDirectory;
        public static string outPutTxt = "";


        public static string GPSRES_Main(string fileName, string dataPath, string oriPath, string savePath)
        {
            outPutTxt = "";
            //// Loading Data
            ArrayList Data = ReadFiles(Path.Combine(dataPath, fileName));

            //// Class to three class
            string[,] CarData = null;
            string[,] GeoData = null;
            string[,] ProData = null;
            int[] dateNum = new int[] { 0, 0, 0 };
            bool check = ThreeCaseProcess(Data, ref CarData, ref GeoData, ref ProData, ref dateNum);

            //// Error type 1 and 2;
            if (check)
            {
                //// If data row error, less one more row data
                return outPutTxt;
            }
            else if (dateNum[0] != dateNum[1] | dateNum[1] != dateNum[2] | dateNum[0] != dateNum[2])
            {
                //// If data number diffrernt, less series data 
                outPutTxt += "資料組數有誤 \r\n";
                outPutTxt += "Cartesian  : " + dateNum[0].ToString() + " 組\r\n";
                outPutTxt += "Geodetic   : " + dateNum[1].ToString() + " 組\r\n";
                outPutTxt += "Projection : " + dateNum[2].ToString() + " 組\r\n";
                return outPutTxt;
            }

            //// if Data no problem, saving to txt and excel files
            string saveName0 = fileName.Replace(".RES", "") + "-Cartesian.txt";
            string saveName1 = fileName.Replace(".RES", "") + "-Geodetic.txt";
            string saveName2 = fileName.Replace(".RES", "") + "-Projection.txt";
            string saveExcelName = fileName.Replace(".RES", ".xlsx");

            SaveToTEXT(savePath, saveName0, dateNum[0], CarData);
            SaveToTEXT(savePath, saveName1, dateNum[1], GeoData);
            SaveToTEXT(savePath, saveName2, dateNum[2], ProData);

            File.Delete(Path.Combine(savePath, saveExcelName));
            File.Copy(Path.Combine(oriPath, "Ori_GPSexcel.xlsx"), Path.Combine(savePath, saveExcelName));
            ExcelClass.ExcelSaveAndRead.Save(strPath: Path.Combine(savePath, saveExcelName), sheetNumber: 1, poRow: 2, poCol: 1, Data: CarData);
            ExcelClass.ExcelSaveAndRead.Save(strPath: Path.Combine(savePath, saveExcelName), sheetNumber: 2, poRow: 2, poCol: 1, Data: GeoData);
            ExcelClass.ExcelSaveAndRead.Save(strPath: Path.Combine(savePath, saveExcelName), sheetNumber: 3, poRow: 2, poCol: 1, Data: ProData);



            //// Sort and Measure Part
            string[,] allData = null;
            bool errCheck = SortAndMeasurements(Data, ref allData);
            //// Error type 3
            if (errCheck)
            {
                outPutTxt = "Measurements(Projection Vectors, Heights) 與 Station->To Station  資料數不同\r\n";
                return outPutTxt;
            }

            //// Calculate
            string[,] sData = new string[allData.GetLength(0), 9];
            processDistance(allData, ref sData);


            ExcelClass.ExcelSaveAndRead.Save(strPath: Path.Combine(savePath, saveExcelName), sheetNumber: 4, poRow: 2, poCol: 1, Data: allData);
            ExcelClass.ExcelSaveAndRead.Save(strPath: Path.Combine(savePath, saveExcelName), sheetNumber: 5, poRow: 2, poCol: 1, Data: sData);



            outPutTxt += "資料處理完畢\r\n";
            return outPutTxt;
        }

        /// <summary>
        /// Part 1 : 處理 水平距與斜距 sheet 5
        /// Part 2 : 比較相同點號的"水平距"與"斜距" 差
        /// </summary>
        /// <param name="allData"></param>
        /// <param name="sData"></param>
        private static bool processDistance(string[,] allData, ref string[,] sData)
        {
            //// Part 1 : Calculate S and s
            string[] number = new string[sData.GetLength(0)];
            string[] point = new string[sData.GetLength(0)];
            string[] sData1 = new string[sData.GetLength(0)];
            string[] sData2 = new string[sData.GetLength(0)];
            for (int i = 0; i < sData.GetLength(0); i++)
            {
                double d1, d2, d3;
                sData[i, 0] = allData[i, 0];
                sData[i, 1] = allData[i, 4];
                d1 = Convert.ToDouble(allData[i, 5]);
                d2 = Convert.ToDouble(allData[i, 8]);
                d3 = Convert.ToDouble(allData[i, 11]);
                sData[i, 2] = Math.Sqrt(d1 * d1 + d2 * d2).ToString();
                sData[i, 3] = Math.Sqrt(d1 * d1 + d2 * d2 + d3 * d3).ToString();
                number[i] = sData[i, 0];
                point[i] = sData[i, 1];
                sData1[i] = sData[i, 2];
                sData2[i] = sData[i, 3];
            }



            //// Part 2 : Comparte the diff S and s of the same station-points
            //// SORT : According to the point Name to sort
            string[] tmpPoint1 = new string[sData.GetLength(0)];
            string[] tmpPoint2 = new string[sData.GetLength(0)];
            point.CopyTo(tmpPoint1, 0);
            point.CopyTo(tmpPoint2, 0);
            Array.Sort(point, number);
            Array.Sort(tmpPoint1, sData1);
            Array.Sort(tmpPoint2, sData2);

            //// Find repeat points
            int lastPo = 0;
            bool flag = true;
            string ind1 = point[0];
            ArrayList repeatPo = new ArrayList();
            for (int i = 1; i < sData.GetLength(0) - 1; i++)
            {
                if (point[i] == ind1)
                {
                    if (flag)
                    {
                        repeatPo.Add(lastPo);
                    }
                    repeatPo.Add(i);
                    flag = false;
                }
                else
                {
                    ind1 = point[i];
                    lastPo = i;
                    flag = true;
                }
            }

            if (repeatPo.Count == 0)
            {
                return true;
            }

            string[] dS = new string[repeatPo.Count];
            string[] ds = new string[repeatPo.Count];
            int po = Convert.ToInt32(repeatPo[0]);
            string refPoint = point[po];
            double refData1 = Convert.ToDouble(sData1[po]);
            double refData2 = Convert.ToDouble(sData2[po]);
            string cmpPoint;
            double cmpData1, cmpData2;
            flag = true;
            for (int ii = 1; ii < repeatPo.Count; ii++)
            {
                po = Convert.ToInt32(repeatPo[ii]);
                cmpPoint = point[po];
                cmpData1 = Convert.ToDouble(sData1[po]);
                cmpData2 = Convert.ToDouble(sData2[po]);

                if (refPoint == cmpPoint)
                {
                    dS[ii] = (refData1 - cmpData1).ToString();
                    ds[ii] = (refData2 - cmpData2).ToString();
                    // Console.WriteLine(dS[ii] + " " + ds[ii]);
                }
                else
                {
                    refData1 = Convert.ToDouble(sData1[po]);
                    refData2 = Convert.ToDouble(sData2[po]);
                    refPoint = point[po];
                }
            }

            string[] errorPoint = new string[repeatPo.Count];
            string[] errorNumber = new string[repeatPo.Count];
            for (int ii = 0; ii < repeatPo.Count; ii++)
            {
                po = Convert.ToInt32(repeatPo[ii]);
                errorPoint[ii] = point[po];
                errorNumber[ii] = number[po];
                //Console.WriteLine(number[po] + " " + point[po] + " " + dS[ii] + " " + ds[ii]);
            }

            //// Indicate to output string[,] 
            for (int i = 0; i < repeatPo.Count; i++)
            {
                sData[i, 5 + 0] = errorNumber[i];
                sData[i, 5 + 1] = errorPoint[i];
                sData[i, 5 + 2] = dS[i];
                sData[i, 5 + 3] = ds[i];
            }
            return false;
        }





        /// <summary>
        /// 處理Sort 與 Measurement 資料, sheet 4
        /// </summary>
        /// <param name="Data"></param>
        /// <param name="allData"></param>
        private static bool SortAndMeasurements(ArrayList Data, ref string[,] allData)
        {
            ArrayList sortData = new ArrayList();
            ArrayList measData = new ArrayList();
            bool OPEN = false;
            bool errCheck = false;
            int check;
            int iter = 0;
            //// Catch two type data
            foreach (string ff in Data)
            {
                //// Catch Sort data
                if (ff.Contains("-->"))
                {
                    sortData.Add(ff);
                }
                else if (ff.Contains("Measurements(Projection Vectors, Heights)"))
                {
                    OPEN = true;
                }

                //// Catch Measurements Data
                if (OPEN)
                {
                    string index = ff;
                    string judge = ff.Substring(0, 5).Trim();
                    if (int.TryParse(judge, out check))
                    {
                        measData.Add(ff);
                        iter++;
                    }
                }

            }

            //// Check two type Data number is Equal ? 
            if (measData.Count != sortData.Count)
            {
                errCheck = true;                                                                                                    //// Error type 3
                return errCheck;
            }

            allData = new string[measData.Count, 14];
            string[] spltData = null;
            for (int i = 0; i < measData.Count; i++)
            {
                //// Part 1 : sort Data
                string index1 = Convert.ToString(sortData[i]);
                string tmpData1 = index1.Substring(0, 6).Trim();                                              //// iter
                string tmpData2 = index1.Substring(6, 3).Trim();                                              //// E or NULL
                string tmpData3 = null;                                                                        //// PATH
                string tmpData4 = index1.Substring(30, index1.Length - 30).Trim();                            //// point --> point
                string tmpData5 = "";                                                                         //// take off head nuber and then save to tmpData5
                spltData = tmpData4.Split(' ');
                for (int p = 1; p < spltData.Length; p++)
                {
                    tmpData5 += spltData[p] + " ";
                }

                string[] pathString = index1.Substring(9, index1.Length - 9).Trim().Split(' ');
                tmpData3 = pathString[0];

                allData[i, 0] = tmpData1;
                allData[i, 2] = tmpData2;
                allData[i, 3] = tmpData3;
                allData[i, 4] = tmpData5.Trim();


                //// Part 2 : get Value
                string index2 = Convert.ToString(measData[i]);                      //// Measurement Data
                string index3 = index2.Substring(6, index2.Length - 6);             //// Take off head number
                spltData = index3.Split(')');                                       //// According to symbol ")" to split string that finally have Three part (N,E,and dh)
                int kk = 0;

                for (int p = 0; p < 3; p++)                                         //// This loop is to control N,E,and dh and each case have Three digit for "coordinate" and "two error"
                {
                    string tmp = spltData[p].Trim();
                    string tmpData7 = tmp.Substring(0, tmp.IndexOf("(")).Trim();
                    string tmpData8 = tmp.Substring(tmp.IndexOf("("), tmp.IndexOf(",") - tmp.IndexOf("(")).Replace("(", "").Trim();
                    string tmpData9 = tmp.Substring(tmp.IndexOf(","), tmp.Length - tmp.IndexOf(",")).Replace(",", "").Trim();
                    allData[i, 5 + kk] = tmpData7;
                    allData[i, 6 + kk] = tmpData8;
                    allData[i, 7 + kk] = tmpData9;
                    kk = kk + 3;
                }
            }
            return errCheck;
        }





        /// <summary>
        /// 截取三種座標系統資料
        /// </summary>
        /// <param name="Data"></param>
        /// <param name="CarData"></param>
        /// <param name="GeoData"></param>
        /// <param name="ProData"></param>
        /// <param name="dateNum"></param>
        private static bool ThreeCaseProcess(ArrayList Data, ref string[,] CarData, ref string[,] GeoData, ref string[,] ProData, ref int[] dateNum)
        {
            string[] titleName = new string[] { "Adjusted Cartesian Coordinates", "Adjusted Geodetic Coordinates", "Adjusted Projection Coordinates", "standard deviation of unit weight" };
            int CASE = 0;
            int kk = 0;
            int[] po = new int[] { 0, 0, 0, 0 };
            //// Find titleName position in oriFile
            foreach (string ff in Data)
            {
                if (CASE <= 3 && ff.Contains(titleName[CASE]))
                {
                    po[CASE] = kk;
                    CASE++;
                }
                else if (CASE > 3)
                {
                    break;
                }
                kk++;
            }
            //string[,] CarData = new string[kk, 7];
            CarData = new string[kk, 7];
            GeoData = new string[kk, 7];
            ProData = new string[kk, 9];
            int check = 0;
            for (int i = 0; i < 3; i++)
            {
                for (int j = po[i]; j < po[i + 1]; j++)
                {
                    string index = Convert.ToString(Data[j]);
                    string judge = index.Substring(0, 5).Trim();
                    if (int.TryParse(judge, out check))
                    {
                        //// Get Data after stationName
                        //// Beacause it have block and garbage string
                        //// using loop to comfirm by check string is contains the char of "("
                        string index2 = null;
                        int wi = 0;
                        bool subCheck = false;
                        while (!subCheck)
                        {
                            wi++;
                            index2 = Convert.ToString(Data[j + wi]);
                            subCheck = index2.Contains("(");
                            if (int.TryParse(index2.Substring(0, 5).Trim(), out check))
                            {
                                outPutTxt += "第" + (j + 1).ToString() + "缺少資料";                                                 //// Error type 1
                                return true;
                            }
                            else if (index2.Contains(titleName[3]) | index2.Contains(titleName[2]) | index2.Contains(titleName[1]))
                            {
                                outPutTxt += "第" + (j + 1).ToString() + "缺少資料";                                                 //// Error type 1
                                return true;
                            }
                        }

                        //// Three Case
                        string stationName = index.Substring(6, 15).Trim();
                        string number = judge;

                        if (i == 0)
                        {
                            int savePo = Convert.ToInt32(number) - 1;
                            string tmpData1 = index2.Substring(0, 27).Trim();
                            string tmpData2 = index2.Substring(27, 24).Trim();
                            string tmpData3 = index2.Substring(51, 21).Trim();
                            CarData[savePo, 0] = stationName;
                            CarData[savePo, 1] = tmpData1.Substring(0, tmpData1.IndexOf("(")).Replace("(", "");
                            CarData[savePo, 2] = tmpData2.Substring(0, tmpData2.IndexOf("(")).Replace("(", "");
                            CarData[savePo, 3] = tmpData3.Substring(0, tmpData3.IndexOf("(")).Replace("(", "");
                            CarData[savePo, 4] = tmpData1.Substring(tmpData1.IndexOf("("), tmpData1.IndexOf(")") - tmpData1.IndexOf("(")).Replace("(", "").Replace(")", "").Trim();
                            CarData[savePo, 5] = tmpData2.Substring(tmpData2.IndexOf("("), tmpData2.IndexOf(")") - tmpData2.IndexOf("(")).Replace("(", "").Replace(")", "").Trim();
                            CarData[savePo, 6] = tmpData3.Substring(tmpData3.IndexOf("("), tmpData3.IndexOf(")") - tmpData3.IndexOf("(")).Replace("(", "").Replace(")", "").Trim();
                            //Console.WriteLine(CarData[savePo, 0] + "  " + CarData[savePo, 1] + " " + CarData[savePo, 2] + " " + CarData[savePo, 3] + " " + CarData[savePo, 4] + " " + CarData[savePo, 5] + " " + CarData[savePo, 6]);

                            dateNum[0]++;
                        }
                        else if (i == 1)
                        {
                            int savePo = Convert.ToInt32(number) - 1;
                            string tmpData1 = index2.Substring(0, 26).Trim();
                            string tmpData2 = index2.Substring(26, 26).Trim();
                            string tmpData3 = index2.Substring(52, 23).Trim();
                            GeoData[savePo, 0] = stationName;
                            GeoData[savePo, 1] = tmpData1.Substring(0, tmpData1.IndexOf("(")).Replace("(", "").Replace("N", "").Trim().Replace(" ", "-").Replace("--", "- ").Replace("-.", " .");
                            GeoData[savePo, 2] = tmpData2.Substring(0, tmpData2.IndexOf("(")).Replace("(", "").Replace("E", "").Trim().Replace(" ", "-").Replace("--", "- ").Replace("-.", " .");
                            GeoData[savePo, 3] = tmpData3.Substring(0, tmpData3.IndexOf("("));
                            GeoData[savePo, 4] = tmpData1.Substring(tmpData1.IndexOf("("), tmpData1.IndexOf(")") - tmpData1.IndexOf("(")).Replace("(", "").Replace(")", "").Trim();
                            GeoData[savePo, 5] = tmpData2.Substring(tmpData2.IndexOf("("), tmpData2.IndexOf(")") - tmpData2.IndexOf("(")).Replace("(", "").Replace(")", "").Trim();
                            GeoData[savePo, 6] = tmpData3.Substring(tmpData3.IndexOf("("), tmpData3.IndexOf(")") - tmpData3.IndexOf("(")).Replace("(", "").Replace(")", "").Trim();
                            // Console.WriteLine(GeoData[savePo, 0] + "  " + GeoData[savePo, 1] + " " + GeoData[savePo, 2] + " " + GeoData[savePo, 3] + " " + GeoData[savePo, 4] + " " + GeoData[savePo, 5] + " " + GeoData[savePo, 6]);

                            dateNum[1]++;
                        }
                        else if (i == 2)
                        {
                            int savePo = Convert.ToInt32(number) - 1;
                            string tmpData0 = index.Substring(index.Length - 12, 12).Trim();
                            string tmpData1 = index2.Substring(0, 26).Trim();
                            string tmpData2 = index2.Substring(26, 24).Trim();
                            string tmpData3 = index2.Substring(52, 18).Trim();
                            string tmpData4 = index2.Substring(index2.Length - 12, 12).Trim();
                            ProData[savePo, 0] = stationName;
                            ProData[savePo, 1] = tmpData1.Substring(0, tmpData1.IndexOf("(")).Replace("(", "");
                            ProData[savePo, 2] = tmpData2.Substring(0, tmpData2.IndexOf("(")).Replace("(", "");
                            ProData[savePo, 3] = tmpData3.Substring(0, tmpData3.IndexOf("(")).Replace("(", "");
                            ProData[savePo, 4] = tmpData1.Substring(tmpData1.IndexOf("("), tmpData1.IndexOf(")") - tmpData1.IndexOf("(")).Replace("(", "").Replace(")", "").Trim();
                            ProData[savePo, 5] = tmpData2.Substring(tmpData2.IndexOf("("), tmpData2.IndexOf(")") - tmpData2.IndexOf("(")).Replace("(", "").Replace(")", "").Trim();
                            ProData[savePo, 6] = tmpData3.Substring(tmpData3.IndexOf("("), tmpData3.IndexOf(")") - tmpData3.IndexOf("(")).Replace("(", "").Replace(")", "").Trim();
                            ProData[savePo, 7] = tmpData0;
                            ProData[savePo, 8] = tmpData4;
                            // Console.WriteLine(tmpData0 + "         " +  tmpData1 + "           " + tmpData2 + "          " + tmpData3 + "         " + tmpData4);
                            //Console.WriteLine(ProData[savePo, 0] + "  " + ProData[savePo, 1] + " " + ProData[savePo, 2] + " " + ProData[savePo, 3] + " " + ProData[savePo, 4] + " " + ProData[savePo, 5] + " " + ProData[savePo, 6] + " " + ProData[savePo, 7] + " " + ProData[savePo, 8]);

                            dateNum[2]++;
                        }
                    }
                }
            }
            return false;
        }

        /// <summary>
        /// 讀檔
        /// </summary>
        /// <param name="dataPath"></param>
        /// <returns></returns>
        private static ArrayList ReadFiles(string dataPath)
        {
            StreamReader sr = new StreamReader(dataPath);
            ArrayList Data = new ArrayList();

            while (sr.Peek() != -1)
            {
                string index = sr.ReadLine();
                Data.Add(index);
            }
            return Data;

        }


        /// <summary>
        /// 截取之資料儲存至txt
        /// </summary>
        /// <param name="savePath"></param>
        /// <param name="saveName"></param>
        /// <param name="dateNum"></param>
        /// <param name="Data"></param>
        private static void SaveToTEXT(string savePath, string saveName, int dateNum, string[,] Data)
        {

            StreamWriter sr = new StreamWriter(Path.Combine(savePath, saveName));
            string[] block = new string[] { "", " ", "  ", "   ", "    ", "     ", "      ", "       ", "        ", "         " };
            int[] maxLen = new int[] { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };
            for (int i = 0; i < dateNum; i++)
            {
                for (int j = 0; j < Data.GetLength(1); j++)
                {
                    maxLen[j] = Data[i, j].Length > maxLen[j] ? Data[i, j].Length : maxLen[j];
                }
            }

            for (int i = 0; i < dateNum; i++)
            {
                string index = "";
                for (int j = 0; j < Data.GetLength(1); j++)
                {
                    index += block[maxLen[j] - Data[i, j].Length] + Data[i, j] + block[6];
                }
                sr.WriteLine(index);
                sr.Flush();
            }
            sr.Close();
        }





    }
}
