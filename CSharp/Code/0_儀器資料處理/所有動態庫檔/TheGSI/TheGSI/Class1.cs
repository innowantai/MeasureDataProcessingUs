using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Collections;
/*
 2018-06-14 14:30

 -- 儀器種類
     初始程式碼為處理"TCRA1201"種類之儀器數據，其餘種類之儀器數據皆須前處理轉換至該之格式
     1. TCRA1201
     2. TC1101  :  (1) data倒數第五位增加小數點 
                   (2) 84.. and 85.. 替換至 87.. and 88..
                   (3) 正倒鏡位置交換

 
 程式概要 : 
 0.處理時會忽略有#符號之列數
 1.讀取GSI檔後將其多餘之0與..符號拿掉儲存較乾淨好閱讀之GSJ文件
 2.將GSI各列資料所需資訊擷取儲存至tmpData(nx8)變數中
        其8行所代表含意依序為 :  0 :[測站1,測點2,皆非-1], 1: 測站名稱, 2: 水平角, 3:垂直角, 4:斜距, 5:覘標高, 6:儀器高, 7:於原始檔GSI之列數 
 3.判斷各測站資料長度是否相同,各測點資料長度是否相同
 4.依照tmpData之第0行資料判斷,將非測站與測點資料拿掉與測站後一行贅資料拿掉轉存至proData
 5.判斷測站與測點是否有重複，重複代表為同一測組之多測回結果，其測站可以忽略拿掉，將拿掉結果儲存至resData中並將測站資訊儲存至statInf中(nx3),其三行依序為 : 0:[測站於resData之列數位置], 2:[測站名稱], 3:[測站於原始檔GSI列數位置]
 6.判斷測點資料是否為四個一組,否則回報錯誤位置
 7.測點資列皆為4個一組後判斷(1)多測回之測站組數名稱是否皆相同 (2)單測回正倒鏡名稱使否相同,否則皆回報錯誤位置
 8.資料皆無誤後排序為excel格式並輸出
 */

namespace TheGSI
{
    public class TheGSI
    {
        public static string oriPath = System.Environment.CurrentDirectory;
        public static string outPutTxt = "";
        public static string[] cName = new string[] { "21.324+", "22.324+", "31..00+", "87..10+", "88..10+" };

        static void Main(string[] args)
        {
            string fileName = "1070102.GSI";
            string dataPath = oriPath;
            string savePath = oriPath;
            string res;
            res = TheGSI_Main(dataPath, savePath, fileName);
            Console.WriteLine(res);
        }

        public static string TheGSI_Main(string dataPath, string savePath, string fileName)
        {
            outPutTxt = "";
            //// 讀檔
            ArrayList data = new ArrayList();
            Read(dataPath, fileName, ref data);

            string ImplementName = (System.String)data[1];
            if (ImplementName != "TCRA1201")
            {
                data = TransDataFormForTC1101(data); 
            } 

            //// 將原始檔儲存至GSJ格式,拿掉不必要之符號,版面乾淨
            SaveToGSJ(savePath, fileName, data);

            //// 擷取所需參數儲存至tmpData
            string[,] tmpData = new string[data.Count, 8];  //// Last is to record ori-position
            CatchInf(data, ref tmpData);

            //// 判斷各測站間和各測點間各自的資料長度是否相同
            outPutTxt = ErrorCheck_0(data, tmpData);
            if (outPutTxt != null) { return outPutTxt; }

            //// 第一階段處理 : 拿掉非觀測資料與測站後第一行贅數
            int[] statPo = new int[data.Count];
            string[,] proData = GetProcessDataStep1(tmpData);

            //// 第二階段處理 : 找出測站位置與名稱,來判斷是否有重複,重複表示可能是同一組與多測回,拿掉所有重複之測站
            string[,] statInf = null;
            string[,] resData = GetProcessDataStep2(proData, ref statInf);

            //// Error1 judge : 判斷測點資料是否皆為四個一組,有則回報錯誤位置
            outPutTxt = ErrorCheck_1(resData, statInf);
            if (outPutTxt != null) { return outPutTxt; }

            //// Error2 judge : 確認資料皆為四個一組後
            ////                (1) 判斷 多測回情況之各組數測點名稱是否有誤,有則回報錯誤位置
            ////                (2) 判斷 單測回情況正倒鏡測點名稱是否有誤,有則回報錯誤位置
            int[] group = new int[statInf.GetLength(0)];
            outPutTxt = ErrorCheck_2(resData, statInf, ref group);
            if (outPutTxt != null) { return outPutTxt; }

            //// 排列至Excel儲存格式
            string[,] excelData = SortToExcelFormat(resData, statInf, group);

            string indexSaveName = (System.String)data[0];
            indexSaveName = indexSaveName.Replace(".GSI", ".xls");
            File.Delete(Path.Combine(savePath, indexSaveName));
            File.Copy(Path.Combine(oriPath, "Ori_excel.xls"), Path.Combine(savePath, indexSaveName));
            ExcelClass.ExcelSaveAndRead.Save(strPath: Path.Combine(savePath, indexSaveName), sheetNumber: 3, poRow: 2, poCol: 1, Data: excelData);

            outPutTxt += "處理完成\r\n";

            return outPutTxt;
        }

        /// <summary>
        /// 轉換儀器格式格式
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        private static ArrayList TransDataFormForTC1101(ArrayList data)
        {

            List<string> tmpData = new List<string>();
            string res = "";
            foreach (string ff in data)
            {
                string[] index = ff.Split(' ');
                //// 增加小數點及更換title
                for (int i = 0; i < index.Length; i++)
                {
                    string tmp = index[i];
                    tmp = tmp.IndexOf("84..") != -1 ? tmp.Replace("84..", "88..") : tmp;
                         //(tmp.IndexOf("85..") != -1 ? tmp.Replace("85..", "88..") : tmp);
                    if (i != 0 & tmp.Trim() != "")
                    {
                        string tmp1 = tmp.Substring(0, tmp.Length - 5);
                        string tmp2 = tmp.Substring(tmp.Length - 5, 5);
                        index[i] = tmp1 + "." + tmp2;
                    }
                    res += index[i] + " ";
                }
                tmpData.Add(res);
                res = "";
            }

            //// 正倒鏡點位互換
            string[] transData = new string[tmpData.Count];
            for (int i = 0; i < tmpData.Count; i++)
            {
                if (i % 6 == 4)
                {
                    transData[i] = tmpData[i];
                    transData[i + 1] = tmpData[i + 3];
                    transData[i + 2] = tmpData[i + 1];
                    transData[i + 3] = tmpData[i + 2];
                }
                else if (i < 2 | i % 6 == 2 | i % 6 == 3)
                {
                    transData[i] = tmpData[i].Trim();
                }
            }

            ArrayList resData = new ArrayList();
            foreach (string ff in transData)
            {
                resData.Add(ff); 
            }

            return resData;

        }

        /// <summary>
        /// 很單純-->讀全部的檔案
        /// </summary>
        /// <param name="dataPath"></param>
        /// <param name="fileName"></param>
        /// <param name="data"></param>
        private static void Read(string dataPath, string fileName, ref ArrayList data)
        {
            StreamReader sr = new StreamReader(Path.Combine(dataPath, fileName));
            while (sr.Peek() != -1)
            {
                string index = sr.ReadLine();
                data.Add(index);
            }
            sr.Close();
        }

        /// <summary>
        /// 截取所需之參數[-1~2,stationName,21,22,31,87,88,oriDataposition]
        /// 將結果儲存至變數tmpData中
        /// </summary>
        /// <param name="data"></param>    原始檔案
        /// <param name="tmpData"></param> OutPut : 總共有8個col,依序為:[-1~2,StationName,21,22,31,87,88,oriDataPosition]
        private static void CatchInf(ArrayList data, ref string[,] tmpData)
        {
            int kk = 0;
            foreach (string ff in data)
            {
                int jj = 2;
                int CASE = 0;
                //// Col-1 : use to save StationName;
                tmpData[kk, 1] = ff[0] == '*' ? takeOffnonMeanData(ff.Substring(ff.IndexOf("+"), ff.IndexOf(" ") - ff.IndexOf("+"))).Replace("+", "").Trim() : "-1";
                //// This loop is to catch information and save into col-2~6, if exist save valur or save "-1"; 
                foreach (string tmpStr in cName)
                {
                    string index = ff; 
                    string index2 = null;
                    int tmpPo = index.IndexOf(tmpStr);
                    bool check = tmpPo != -1 ? true : false;
                    if (check && !index.Contains("#"))
                    {
                        index = index.Substring(tmpPo, index.Length - index.IndexOf(tmpStr));
                        int po2 = index.IndexOf(" ") == -1 ? index.Length : index.IndexOf(" ");
                        index2 = index.Substring(tmpStr.Length, po2 - tmpStr.Length);
                        CASE++;
                    }
                    else
                    {
                        index2 = "-1";
                    } 
                    tmpData[kk, jj] = Convert.ToDouble(index2).ToString();
                    jj++;
                } 
                //// Col-0 : use to tag CASE(-1 1 2) 
                tmpData[kk, 0] = CASE == 0 ? "-1" : (CASE == 1 ? "1" : "2");
                //// Col-7 : use to record oriDataPosition
                tmpData[kk, 7] = kk.ToString();
                kk++;
            }
            //for (int i = 0; i < kk; i++)
            //{
            //    Console.WriteLine(i + " : " + tmpData[i, 0] + " " + tmpData[i, 1] + " " + tmpData[i, 2] + " " + tmpData[i, 3] + " " + tmpData[i, 4] + " " + tmpData[i, 5] + " " + tmpData[i, 6]);
            //}
        }

        /// <summary>
        /// 將第一階段擷取之tmpData資料進一步篩選，拿掉非"量測資料"與"測站後一行之資料"
        /// 紀錄所有測站位置，用來判斷資料長度是否皆相同
        /// </summary> 
        /// <param name="tmpData"></param>
        /// <returns> proData </returns> 輸出整理後之結果
        private static string[,] GetProcessDataStep1(string[,] tmpData)
        {
            int kk = 0;
            bool jump = false;
            string[,] proData = new string[tmpData.GetLength(0), 8];
            for (int i = 0; i < proData.GetLength(0); i++)
            {
                if (tmpData[i, 0] == "-1" | jump)
                {
                    jump = false;
                    continue;
                }
                else
                {
                    for (int j = 0; j < proData.GetLength(1); j++)
                    {
                        proData[kk, j] = tmpData[i, j];
                    }
                    kk++;
                    jump = false;
                }
                jump = tmpData[i, 0] == "1" ? true : false;
            }

            //for (int i = 0; i < proData.GetLength(0); i++)
            //{
            //    Console.WriteLine(i + " : " + proData[i, 0] + " " + proData[i, 1] + " " + proData[i, 2] + " " + proData[i, 3] + " " + proData[i, 4] + " " + proData[i, 5] + " " + proData[i, 6] + " " + proData[i, 7]);
            //}
            return proData;
        }

        /// <summary>
        /// 找出測站位置與名稱,來判斷是否有重複,重複表示可能是同一組與多測回
        /// 拿掉所有重複不要之測站後，儲存新變數並記錄其位置輸出
        /// </summary>
        /// <param name="proData"></param>
        /// <param name="statInf"></param>
        /// <returns>resData　statInf</returns>　分別為 : 量測資料 與 測站資訊(Col [0,1,2] = [測站於resData位置,測站名稱,測站於原始GSI檔之位置])
        private static string[,] GetProcessDataStep2(string[,] proData, ref string[,] statInf)
        {

            //// 找出測站位置與名稱,用來判斷是否有重複,重複表示可能是同一組與多測回
            string[,] stationPo = new string[proData.GetLength(0), 3];
            int num = 0;
            for (int i = 0; i < proData.GetLength(0); i++)
            {
                if (proData[i, 0] == "1")
                {
                    stationPo[num, 0] = i.ToString();
                    stationPo[num, 1] = proData[i, 1];
                    stationPo[num, 2] = proData[i, 7];
                    //Console.WriteLine(stationPo[num, 0] + "," + stationPo[num, 1] + "," + stationPo[num, 2]);
                    num++;
                }
            }

            //// 找出重複測站,將重複多餘之測站名稱於proData變數中消除(Null)
            //// 以測站名稱與隔一個點號名稱進行判斷是否為重複測站,同一組測點之多測回應為相同，否則判斷為獨立測回
            //string index = stationPo[0, 1];
            int po = Convert.ToInt32(stationPo[0, 0]);
            string index = proData[po, 1] + proData[po + 1, 1];
            for (int i = 1; i < num; i++)
            {
                po = Convert.ToInt32(stationPo[i, 0]);
                if (proData[po, 1] + proData[po + 1, 1] == index)
                {
                    proData[po, 1] = null;
                    //Console.WriteLine(stationPo[i, 0] + "," + stationPo[i, 1] + "," + stationPo[i, 2]);
                }
                else
                {
                    //index = stationPo[i, 1];
                    index = proData[po, 1] + proData[po + 1, 1];
                }
                //Console.WriteLine(stationPo[i, 0] + "," + stationPo[i, 1] + "," + stationPo[i, 2]);
            }

            //// 重新整理proData : 將不要之測站拿掉
            string[,] resData_ = new string[proData.GetLength(0), 8];
            int kk = 0;
            for (int i = 0; i < resData_.GetLength(0); i++)
            {
                if (proData[i, 1] != null)
                {
                    for (int j = 0; j < resData_.GetLength(1); j++)
                    {
                        resData_[kk, j] = proData[i, j];
                    }
                    kk++;
                }
            }

            //// 拿掉resData_ 空白部分並記錄測站位置與數量
            string[,] resData = new string[kk, 8];
            string[,] statInf_ = new string[kk, 3];     ////Col [0,1,2] = [po at resData,stationName,po at orifile]
            int num2 = 0;
            bool open = false;
            for (int i = 0; i < kk; i++)
            {
                if (resData_[i, 1] != null)
                {
                    for (int j = 0; j < resData_.GetLength(1); j++)
                    {
                        resData[i, j] = resData_[i, j];
                    }
                }
                //// Judge is the station ?
                open = resData[i, 0] == "1" ? true : false;
                if (open)
                {
                    statInf_[num2, 0] = i.ToString();
                    statInf_[num2, 1] = resData[i, 1];
                    statInf_[num2, 2] = resData[i, 7];
                    // Console.WriteLine(i + " : " + statInf_[num2, 0]+ " " + statInf_[num2, 1] + " " + statInf_[num2, 2]);
                    num2++;
                }
            }

            //// 拿掉 statInf_ 空白部分
            statInf = new string[num2, 3];
            for (int i = 0; i < num2; i++)
            {
                statInf[i, 0] = statInf_[i, 0];
                statInf[i, 1] = statInf_[i, 1];
                statInf[i, 2] = statInf_[i, 2];
            }

            //for (int i = 0; i < resData.GetLength(0); i++)
            //    Console.WriteLine(i + " : " + resData[i, 0] + " " + resData[i, 1] + " " + resData[i, 2] + " " + resData[i, 3] + " " + resData[i, 4] + " " + resData[i, 5] + " " + resData[i, 6] + " " + resData[i, 7]);

            return resData;
        }

        /// <summary>
        /// 判斷各測站間的資料長度是否相同與判斷各測點間的資料長度是否相同
        /// </summary>
        /// <param name="data"></param>
        /// <param name="tmpData"></param>
        /// <returns></returns>
        private static string ErrorCheck_0(ArrayList data, string[,] tmpData)
        {
            //// errStr 為回報字串，若最後都無錯誤則回傳空值
            string errStr = null;

            int lastkk1 = 0;
            int lastkk2 = 0;
            int kk = 0;
            int[] errPo = new int[tmpData.GetLength(0)];
            for (int i = 0; i < tmpData.GetLength(0); i++)
            {
                string index = (System.String)data[i];
                if (tmpData[i, 0] == "1")
                {
                    lastkk1 = lastkk1 == 0 ? index.Length : lastkk1;
                    if (lastkk1 != 0 && index.Length != lastkk1)
                    {
                        errPo[kk] = i + 1;
                        kk++;
                    }
                }
                else if (tmpData[i, 0] == "2")
                {
                    lastkk2 = lastkk2 == 0 ? index.Length : lastkk2;
                    if (lastkk2 != 0 && index.Length != lastkk2)
                    {
                        errPo[kk] = i + 1;
                        kk++;
                    }
                }
              //errStr += tmpData[i, 0] + "\n";
            }

            if (kk == 0) { return errStr; }

            errStr = "第 ";
            foreach (int ii in errPo)
            {
                if (ii != 0)
                {
                    errStr += ii.ToString() + ", ";
                }
            }
            errStr += "列資料長度有誤";
            return errStr;
        }

        /// <summary>
        /// 判斷量測組數是否為4個一組
        /// 若不是回報錯誤並給予於原始檔之錯誤列數位置
        /// </summary>
        /// <param name="proData2"></param>
        /// <param name="statInf"></param>
        /// <returns> errStr </returns>    有誤回傳錯誤列數否則空白
        private static string ErrorCheck_1(string[,] proData2, string[,] statInf)
        {
            //// errStr 為回報字串，若最後都無錯誤則回傳空值
            string errStr = null;

            //// 找出測站間的資料組數(正常必須為4的倍數)
            int[] chNumber = new int[statInf.GetLength(0)];
            for (int i = 0; i < statInf.GetLength(0); i++)
            {
                int po1 = Convert.ToInt32(statInf[i, 0]);
                int po2 = i == statInf.GetLength(0) - 1 ? proData2.GetLength(0) : Convert.ToInt32(statInf[i + 1, 0]);
                chNumber[i] = po2 - po1 - 1;
                //Console.WriteLine(po2 + "," + po1 + "," + chNumber[i]);
            }

            //// 判斷測站組數是否為4的倍數，否則回傳非4倍數之資料於原始GSI檔位置
            for (int i = 0; i < chNumber.GetLength(0); i++)
            {
                int num = chNumber[i];
                if (num % 4 != 0)
                {
                    int po1 = Convert.ToInt32(statInf[i, 0]);
                    int po2 = i == statInf.GetLength(0) - 1 ? proData2.GetLength(0) : Convert.ToInt32(statInf[i + 1, 0]);
                    for (int j = po1 + 1; j < po2 - 4; j = j + 4)
                    {
                        //// Try error : 最後邊界超出去為獨立一個Case
                        try
                        {
                            string index1 = proData2[j, 1] + proData2[j + 1, 1] + proData2[j + 2, 1] + proData2[j + 3, 1];
                            string index2 = proData2[j + 4, 1] + proData2[j + 5, 1] + proData2[j + 6, 1] + proData2[j + 7, 1];
                            if (index1 != index2)
                            {
                                string tmp = null;
                                errStr = "第 ";
                                for (int ll = 0; ll < 8; ll++)
                                {
                                    tmp = (Convert.ToInt32(proData2[j + ll, 7]) + 1).ToString();
                                    errStr += tmp + ", ";
                                }
                                errStr += " 列之測站組數有誤(不為5個為一組)";
                                return errStr;
                            }
                        }
                        catch (Exception)
                        {
                            errStr += "第" + (Convert.ToInt32(statInf[i, 2]) + 1).ToString() + "列開始之測站組數有誤(不為5個為一組)\r\n";
                            return errStr;
                        }
                    }
                    errStr += "第" + (Convert.ToInt32(statInf[i, 2]) + 1).ToString() + "列開始之測站組數有誤(不為5個為一組)\r\n";
                    return errStr;
                }
            }
            return errStr;
        }

        /// <summary>
        /// 於ErrorCheck_1以確認資料組數無誤皆為4個一組,判斷每四個一組是否皆為相同點號名
        /// 判斷測站名稱是否有與測點名稱相同
        /// </summary>
        /// <param name="proData2"></param>
        /// <param name="statInf"></param>
        /// <returns> errStr group</returns>    有誤回傳錯誤列數否則空白, group :記錄每測站測回數
        private static string ErrorCheck_2(string[,] proData2, string[,] statInf, ref int[] group)
        {

            //for (int i = 0; i < proData2.GetLength(0); i++)
            //{
            //    Console.WriteLine(proData2[i, 7] + "," + proData2[i, 0] + "," + proData2[i, 1]);
            //}

            string errStr = null;
            for (int i = 0; i < statInf.GetLength(0); i++)
            {
                //// 判斷多測回之測點名稱是否有相同,否則回報不同位置
                int num = 1;
                int po1 = Convert.ToInt32(statInf[i, 0]);
                int po2 = i == statInf.GetLength(0) - 1 ? proData2.GetLength(0) : Convert.ToInt32(statInf[i + 1, 0]);
                string statName = proData2[po1, 1];
                for (int j = po1 + 1; j < po2 - 4; j = j + 4)
                {
                    string index1 = proData2[j, 1] + proData2[j + 1, 1] + proData2[j + 2, 1] + proData2[j + 3, 1];
                    string index2 = proData2[j + 4, 1] + proData2[j + 5, 1] + proData2[j + 6, 1] + proData2[j + 7, 1];
                    if (index1 != index2)
                    {
                        string tmp = null;
                        errStr = "第 ";
                        for (int ll = 0; ll < 8; ll++)
                        {
                            tmp = (Convert.ToInt32(proData2[j + ll, 7]) + 1).ToString();
                            errStr = ll == 4 ? errStr + "與 " : errStr;
                            errStr += tmp + ", ";
                        }
                        errStr += " 列之點號名稱有誤(兩組名稱應相同)\r\n";
                        // Console.WriteLine(index1);
                        // Console.WriteLine(index2);
                        return errStr;
                    }
                    else
                    {
                        num++;
                    }
                    //// 判斷測站名稱是否有與測點名稱相同
                    bool check1 = false;
                    bool check2 = false;
                    for (int kkk = 0; kkk < 3; kkk++)
                    {
                        if (proData2[j + kkk, 1].Contains(statName) && proData2[j + kkk, 1].Length == statName.Length)
                        {
                            check1 = true;
                            break;
                        }
                        if (proData2[j + 4 + kkk, 1].Contains(statName) && proData2[j + 4 + kkk, 1].Length == statName.Length)
                        {
                            check2 = true;
                            break;
                        }
                    }
                    if (check1 | check2)
                    {
                        string tmp = null;
                        errStr = "第 ";
                        for (int ll = 0; ll < 8; ll++)
                        {
                            tmp = (Convert.ToInt32(proData2[j + ll, 7]) + 1).ToString();
                            errStr += tmp + ", ";
                        }
                        errStr += " 列 測站名稱與測點名稱相同\r\n";
                        return errStr;
                    }
                }
                group[i] = num;
                //// num == 1時表示為單側回，檢驗單側回之測點名稱是否兩兩相同，否則回報錯誤
                if (num == 1)
                {
                    int j = po1 + 1;
                    string index1 = proData2[j, 1] + proData2[j + 2, 1];
                    string index2 = proData2[j + 1, 1] + proData2[j + 3, 1];
                    statName = proData2[po1, 1];
                    if (index1 != index2)
                    {
                        string tmp = null;
                        errStr = "第 ";
                        for (int ll = 0; ll < 4; ll++)
                        {
                            tmp = (Convert.ToInt32(proData2[j + ll, 7]) + 1).ToString();
                            errStr += tmp + ", ";
                        }
                        Console.WriteLine(index1);
                        Console.WriteLine(index2);
                        errStr += "列之測站名稱有誤,沒有兩兩成對\r\n";
                        return errStr;
                    }

                    //// 判斷測站名稱是否有與測點名稱相同
                    if (index1.Contains(statName) | index2.Contains(statName))
                    {
                        string tmp = null;
                        for (int ll = 0; ll < 4; ll++)
                        {
                            tmp = (Convert.ToInt32(proData2[j + ll, 7]) + 1).ToString();
                            errStr += tmp + ", ";
                        }
                        return errStr + "列 測站名稱與測點名稱相同\r\n";
                    }
                }
            }

            return errStr;
        }

        /// <summary>
        /// 依照group之測回數將resData資料排序成excel格式
        /// 總共有四個迴圈:
        /// (1) 測站數目
        /// (2) 每測站之測回數
        /// (3) 正倒鏡,原始檔四個順序為 (A+)(A-)(B+)(B-),excel格式順序為(A+)(B+)(B-)(A-) -->[0,2,3,1],
        /// (4) 水平,垂直,斜距 
        /// </summary>
        /// <param name="resData"></param>
        /// <param name="statInf"></param>
        /// <param name="group"></param>
        /// <returns> excelData </returns> 
        private static string[,] SortToExcelFormat(string[,] resData, string[,] statInf, int[] group)
        {
            string[,] excelData = new string[statInf.GetLength(0), 220];
            int[] catIter = new int[4] { 0, 2, 3, 1 };
            int kk = 0;
            int po1 = 0;
            for (int i = 0; i < statInf.GetLength(0); i++)
            {
                po1 = Convert.ToInt32(statInf[i, 0]) + 1;
                kk = 6;
                for (int num = 0; num < group[i]; num++)
                {
                    for (int j = 0; j < 4; j++)
                    {
                        int cat = catIter[j];
                        for (int pp = 2; pp < 5; pp++)
                        {
                            excelData[i, kk] = resData[po1 + cat, pp];
                            kk++;
                        }
                    }
                    po1 += 4;
                }

                //// 測站名稱與斜距
                kk = 0;
                po1 = Convert.ToInt32(statInf[i, 0]);
                for (int pp = 0; pp < 6; pp = pp + 2)
                {
                    excelData[i, pp] = resData[po1 + kk, 1];
                    excelData[i, pp + 1] = resData[po1 + kk, 6];
                    kk = kk + 2;
                    //Console.WriteLine(resData[po1 + kk, 1]);
                }
            }

            return excelData;
        }

        /// <summary>
        /// 儲存較乾淨的儲存GSJ檔案
        /// </summary>
        /// <param name="savePath"></param>
        /// <param name="fileName"></param>
        /// <param name="data"></param>
        private static void SaveToGSJ(string savePath, string fileName, ArrayList data)
        {
            StreamWriter sw = new StreamWriter(Path.Combine(savePath, fileName.Replace(".GSI", ".GSJ")));
            foreach (string ff in data)
            {
                sw.WriteLine(takeOffnonMeanData(ff));
                sw.Flush();
            }
            sw.Close();
        }

        /// <summary>
        /// 將空白還有其他亂的沒用的符號拿掉儲存GSJ檔案
        /// </summary>
        /// <param name="savePath"></param>
        /// <param name="fileName"></param>
        /// <param name="data"></param>
        private static string takeOffnonMeanData(string data)
        {
            string[] replaceCase = new string[] { "..10", "..16", "..00" };

            string index = "";
            bool open = false;
            foreach (char ss in data)
            {
                if (open)
                {
                    if (ss == '0')
                    {
                        index += ' ';
                    }
                    else
                    {
                        index += ss;
                        open = false;
                    }
                }
                else
                {
                    index += ss;
                }

                if (ss == '+' | ss == '-')
                {
                    open = true;
                }
            }

            foreach (string ff in replaceCase)
            {
                index = index.Replace(ff, "____");
            }

            index = index.Replace("..", "__");
            return index;
        }
    }
}
