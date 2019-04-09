using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace _1_WriteCodeVersion_2
{
    public class TransCoordinateSystem
    {
        private List<double> Parameter_N = new List<double>() {2800000,2750000,2700000,
                                                               2650000,2600000,2550000,2500000,
                                                               2450000,2400000    };
        private List<double> Parameter_E = new List<double>() { 90000, 170000, 250000, 330000, 410000 };

        private string[,] MappingArray = new string[,]
        {
            {"XX"   , "XX"  , "XX"  , "XX" },
            {"XX"   , "A"  , "B"  , "C" },
            {"XX"   , "D"   , "E"   , "F"},
            {"XX"   , "G"   , "H"   , "I"},
            {"J"   , "K"   , "L"   , "XX"},
            {"M"   , "N"   , "O"   , "XX"},
            {"P"   , "Q"   , "R"   , "XX"},
            {"XX"   , "T"   , "U "   , "XX"},
            {"XX"   , "V"   , "XX"   , "XX"},
        };

        private List<string> QxQy = new List<string>() { "A", "B", "C", "D", "E", "F", "G", "H", };

        private string[,] Title1 = new string[,] { { "點號", "N", "E", "正高", "橢球高", "N", "E", "正高", "橢球高" } };
        private string[,] Title2 = new string[,] { { "點號", "N", "E", "正高", "橢球高", "點號", "N", "E", "正高" } };
        private string Path_Desk = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

        public TransCoordinateSystem()
        {


        }

        public void Main_start()
        {
            //string savePath = Path.Combine(this.Path_Desk, "res.xlsx");
            //string path = Path.Combine(this.Path_Desk, "新增資料夾", "圖號坐標轉換_.xlsx");
            //string[,] OriData = ReadExcelData(path,out string[,] AllDatas);

            //EXCEL excel = new EXCEL("");
            //excel.Save_To(savePath, "OriData", 1, 1, CombineArray(this.Title1, AllDatas));

            //for (int i = 0; i < OriData.GetLength(0); i++)
            //{
            //    string res = Main_TWD97toTWD67(ref OriData, i);
            //    Console.WriteLine(res);
            //}

            //string[,] newTitle = new string[,] { { "TWD97", "", "", "", "", "TWD67", } };
            //excel.Save_To(savePath, "TWD97轉TWD67", 1, 1, CombineArray(CombineArray(newTitle, this.Title2), OriData));
            //excel.close();
            List<double> data = new List<double>() {2546852.303,	178343.647};
           var res =  Main_TWD97toTWD67_debug(data[0], data[1], 0);
            Console.WriteLine(res);
        }

        public string Main_TWD97toTWD67_debug(double input_N, double input_E, double HH)
        { 

            double Trans_N = input_N + 248.6 - 0.00001549 * input_N - 0.000006521 * input_E;
            double Trans_E = input_E - 807.8 - 0.00001549 * input_E - 0.000006521 * input_N;

            int index_E = FindIndexof_E(Trans_E);
            int index_N = FindIndexof_N(Trans_N);
            if (index_E == -1 || index_N == -1) return "-1" + Trans_N + "," + Trans_E;

            string MapRes = MappingArray[index_N, index_E];
            double N1 = Math.Floor((Trans_E - this.Parameter_E[index_E]) / 800);
            double N2 = Math.Floor((Trans_N - this.Parameter_N[index_N]) / 500);
            double Rem_E = (Trans_E - Parameter_E[index_E]) % 800;
            double Rem_N = (Trans_N - Parameter_N[index_N]) % 500;
            string N3 = QxQy[Convert.ToInt32(Math.Floor(Rem_E / 100))];
            string N4 = QxQy[Convert.ToInt32(Math.Floor(Rem_N / 100))];
            double R3 = Math.Round(Rem_E % 100);
            double R4 = Math.Round(Rem_N % 100);
            double N5 = Math.Floor(R3 / 10);
            double N6 = Math.Floor(R4 / 10);
            double N7 = Math.Round(R3 % 10);
            double N8 = Math.Round(R4 % 10);

            string RESULTs = MapRes + N1.ToString() + N2.ToString() + N3 + N4 + N5.ToString() + N6.ToString() + N7.ToString() + N8.ToString();
             
            return RESULTs + "," + Trans_N.ToString() + "," + Trans_E.ToString() + "," + HH.ToString();

        }



        public string Main_TWD97toTWD67(ref string[,] resData, int index)
        {
            double[,] NEs = GetNE(resData, index);
            double input_N = NEs[0, 0];
            double input_E = NEs[0, 1];
            double HH = NEs[0, 2];

            double Trans_N = input_N + 248.6 - 0.00001549 * input_N - 0.000006521 * input_E;
            double Trans_E = input_E - 807.8 - 0.00001549 * input_E - 0.000006521 * input_N;

            int index_E = FindIndexof_E(Trans_E);
            int index_N = FindIndexof_N(Trans_N);
            if (index_E == -1 || index_N == -1) return "-1" + Trans_N + "," + Trans_E;

            string MapRes = MappingArray[index_N, index_E];
            double N1 = Math.Floor((Trans_E - this.Parameter_E[index_E]) / 800);
            double N2 = Math.Floor((Trans_N - this.Parameter_N[index_N]) / 500);
            double Rem_E = (Trans_E - Parameter_E[index_E]) % 800;
            double Rem_N = (Trans_N - Parameter_N[index_N]) % 500;
            string N3 = QxQy[Convert.ToInt32(Math.Floor(Rem_E / 100))];
            string N4 = QxQy[Convert.ToInt32(Math.Floor(Rem_N / 100))];
            double R3 = Math.Round(Rem_E % 100);
            double R4 = Math.Round(Rem_N % 100);
            double N5 = Math.Floor(R3 / 10);
            double N6 = Math.Floor(R4 / 10);
            double N7 = Math.Round(R3 % 10);
            double N8 = Math.Round(R4 % 10);

            string RESULTs = MapRes + N1.ToString() + N2.ToString() + N3 + N4 + N5.ToString() + N6.ToString() + N7.ToString() + N8.ToString();
            resData[index, 5] = RESULTs;
            resData[index, 6] = Trans_N.ToString();
            resData[index, 7] = Trans_E.ToString();
            resData[index, 8] = HH.ToString();
            return RESULTs + "," + Trans_N.ToString() + "," + Trans_E.ToString() + "," + HH.ToString();

        }


        private double[,] GetNE(string[,] Data, int index)
        {
            double[,] res = new double[1, 3];
            int i = index;
            if (Data[i, 1].Trim() != "") res[0, 0] = Convert.ToDouble(Data[i, 1]);
            if (Data[i, 2].Trim() != "") res[0, 1] = Convert.ToDouble(Data[i, 2]);
            if (Data[i, 4].Trim() != "") res[0, 2] = Convert.ToDouble(Data[i, 3]);

            return res;
        }

        private int FindIndexof_N(double Trans_N)
        {
            ; for (int i = 0; i < this.Parameter_N.Count; i++)
            {
                if (this.Parameter_N[i] - Trans_N <= 0)
                {
                    return i;
                }
            }

            return -1;
        }

        private int FindIndexof_E(double Trans_E)
        {
            for (int i = 0; i < this.Parameter_E.Count; i++)
            {
                if (this.Parameter_E[i] - Trans_E >= 0)
                {
                    return i-1;
                }

            }
            return -1;
        }
         

        private string[,] CombineArray(string[,] Data1, string[,] Data2)
        {
            string[,] Result = new string[Data1.GetLength(0) + Data2.GetLength(0), 10];
            for (int i = 0; i < Data1.GetLength(0); i++)
            {
                for (int j = 0; j < Data1.GetLength(1); j++)
                {
                    Result[i, j] = Data1[i, j];
                }
            }

            for (int i = Data1.GetLength(0); i < Data2.GetLength(0) + Data1.GetLength(0); i++)
            {
                for (int j = 0; j < Data2.GetLength(1); j++)
                {
                    Result[i, j] = Data2[i - Data1.GetLength(0), j];
                }
            }

            return Result;
        }


        private string[,] ReadExcelData(string path,out string[,] data)
        {
            EXCEL excel = new EXCEL(path);
             data = excel.GetDataBySheetNumber(2, 1, 1);
            excel.close();
            string[,] resData = new string[data.GetLength(0), 9];
            for (int i = 0; i < resData.GetLength(0); i++)
            {
                resData[i, 0] = data[i, 1];
                resData[i, 1] = data[i, 5];
                resData[i, 2] = data[i, 6];
                resData[i, 3] = data[i, 7];
                resData[i, 4] = data[i, 8];
            }
            return resData;
        }



    }
}
