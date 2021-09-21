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
    public class TransCoordinateSystem
    {
        private List<double> Parameter_N = new List<double>() {2800000,2750000,2700000,
                                                               2650000,2600000,2550000,2500000,
                                                               2450000,2400000    };
        private List<double> Parameter_E = new List<double>() { 90000, 170000, 250000, 330000, 410000 };

        private string[,] MappingArray = new string[,]
        {
            {""   , ""  , ""  , "" },
            {"※"   , "A"  , "B"  , "C" },
            {"※"   , "D"   , "E"   , "F"},
            {"※"   , "G"   , "H"   , "I"},
            {"J"   , "K"   , "L"   , "※"},
            {"M"   , "N"   , "O"   , "※"},
            {"P"   , "Q"   , "R"   , "※"},
            {"※"   , "T"   , "U"   , "※"},
            {"※"   , "V"   , "※"   , "※"},
        };

        private List<string> QxQy = new List<string>() { "A", "B", "C", "D", "E", "F", "G", "H", };

        private string[,] Title1 = new string[,] { { "點號", "N", "E", "正高", "橢球高", "N", "E", "正高", "橢球高" } };
        private string[,] Title2 = new string[,] { { "點號", "N", "E", "正高", "橢球高", "點號", "N", "E", "正高" } };
        private string[,] newTitle = new string[,] { { "TWD97", "", "", "", "", "TWD67", } };
        private string[,] ResultTopTitle = new string[,] {
            { "配合辦理管線定位測量、管線施作及圖資更新維護作業明細表"  },
            { "ＤＣＩＳ:" },
            { "施工號碼:" },  };
        private string[,] ResultBtnTitle = new string[,] {
            {"" ,"","" ,"" ,"","","","",""},
            {"" ,"","" ,"" ,"","","","",""},
            {"" ,"","" ,"測量技術士簽名:" ,"","","","","" },
            {"" ,"","" ,"    ☐符合                 ☐未符合" ,"","","","",""},
            {"" ,"","" ,"" ,"","","","",""},
            {"" ,"","" ,"經辦人:","","","課長:","","" },
            {"" ,"","" ,"" ,"","","","",""}, };


        public void Main_start(string FileFullPath, string SavePath, string SaveFileName)
        {
            string savePath = Path.Combine(SavePath, SaveFileName);
            string[,] OriData = ReadExcelData(FileFullPath, out string[,] AllDatas, out string SheetName);

            if (null == OriData) return;

            EXCEL excel = new EXCEL(savePath);
            excel.Save(SheetName, 1, 1, CombineArray(this.Title1, AllDatas));

            for (int i = 0; i < OriData.GetLength(0); i++) Main_TWD97toTWD67(ref OriData, i);

            string[,] res_1 = CombineArray(CombineArray(this.newTitle, this.Title2), OriData);
            string[,] res_2 = CombineArray(this.ResultTopTitle, res_1);
            string[,] res_3 = CombineArray(res_2, this.ResultBtnTitle);
            //excel.Save("TWD97轉TWD67", 1, 1, CombineArray(CombineArray(this.newTitle, this.Title2), OriData));
            excel.Save_ChangeFormat("TWD97轉TWD67", 1, 1, res_3);
            excel.close();
        }



        private string Main_TWD97toTWD67(ref string[,] resData, int index)
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
            R3 = R3 >= 100 ? (R3 - 100) : R3;
            double N5 = Math.Floor(R3 / 10);
            double N6 = Math.Floor(R4 / 10);
            double N7 = Math.Round(R3 % 10);
            double N8 = Math.Round(R4 % 10);

            string newN1 = N1.ToString().Length == 1 ? N1.ToString().PadLeft(2, '0') : N1.ToString().PadLeft(2, '0');
            string newN2 = N2.ToString().Length == 1 ? N2.ToString().PadLeft(2, '0') : N2.ToString().PadLeft(2, '0');
            string RESULTs = MapRes + newN1 + newN2 + N3 + N4 + N5.ToString() + N6.ToString() + N7.ToString() + N8.ToString();
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
                    return i - 1;
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


        private string[,] ReadExcelData(string path, out string[,] data, out string sheetName)
        {
            EXCEL excel = new EXCEL(path);
            data = excel.GetDataBySheetNumber(2, 1, 1);
            sheetName = excel.sheets[0];
            excel.close();
            // foreach (string item in excel.sheets) if (sheetName.Count() > 1 && item.Contains("TWD97轉TWD67")) return null;


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
