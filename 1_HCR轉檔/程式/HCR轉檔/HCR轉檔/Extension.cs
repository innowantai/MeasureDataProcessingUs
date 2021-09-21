using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using System.Diagnostics;
using System.Threading.Tasks;
using System.Collections.Concurrent;
using System.Runtime.InteropServices;
using System.IO;

namespace HCR轉檔
{
    public static class ExtensionOperation
    {
        public static RowData Distint(this RowData rd)
        {
            if (rd.data.Length > 2)
            {
                string[] tmp = new string[2];
                tmp[0] = rd.data[0];
                tmp[1] = rd.data[1];
                return new RowData(tmp);
            }
            return rd;
        }

        public static string ToOutFormat(this RowData rd)
        {
            string res = "";
            foreach (string d in rd.data)
            {
                string item = null == d ? "" : d;
                bool IsDouble = double.TryParse(item, out double dd);
                string tmp = IsDouble ? dd.ToString() : item;
                res = res + tmp.ToBlockData();
            }

            return res;
        }


        public static string ToBlockData(this string data)
        {
            string tmp = "";

            for (int i = 0; i < 10 - data.Length; i++) tmp = tmp + " ";
            data = tmp + data;

            return data;

        }

        public static string[] SplieToArray(this string data)
        {
            List<string> dd = new List<string>();
            foreach (string item in data.Split(' '))
            {
                if (item.Trim() == "") continue;

                dd.Add(item);
            }
            return dd.ToArray();
        }


        public static List<CROSS_DATA> ToCROSS_DATA(this List<RowData> data)
        {
            List<int> indexes = new List<int>();
            for (int i = 0; i < data.Count; i++)
            {
                if (data[i].data.Contains("CROSS"))
                {
                    indexes.Add(i);
                }
            }
            indexes.Add(data.Count);

            List<CROSS_DATA> CRs = new List<CROSS_DATA>();
            for (int i = 0; i < indexes.Count - 1; i++)
            {
                List<RowData> rds = new List<RowData>();
                for (int j = indexes[i]; j < indexes[i + 1]; j++)
                {
                    rds.Add(data[j]);
                }

                CRs.Add(new CROSS_DATA(rds));

            }

            return CRs;
        }



        public static RowData ToRowData(this IEnumerable<string> data)
        {
            return new RowData(data.ToArray());
        }
        public static ColData ToColData(this IEnumerable<string> data)
        {
            return new ColData(data.ToArray());
        }

        public static bool IsEmpty(this ColData data)
        {
            foreach (string item in data.data)
            {
                if (item.Trim() != "") return false;
            }

            return true;
        }

        public static bool IsEmpty(this RowData data)
        {
            foreach (string item in data.data)
            {
                if (item.Trim() != "") return false;
            }

            return true;
        }

        public static ColData GenerateConstantValueColData(this string dd, int num)
        {
            string[] res = new string[num];
            for (int i = 0; i < num; i++) res[i] = dd;
            return new ColData(res);
        }

        public static RowData GenerateConstantValueRowData(this string dd, int num)
        {
            string[] res = new string[num];
            for (int i = 0; i < num; i++) res[i] = dd;
            return new RowData(res);
        }


        public static List<RowData> ToRowData(this string[,] data)
        {
            List<RowData> res = new List<RowData>();
            int num = data.GetLength(1);
            for (int i = 0; i < data.GetLength(0); i++)
            {
                List<string> dd = new List<string>();
                for (int j = 0; j < data.GetLength(1); j++)
                {
                    dd.Add(data[i, j]);
                }

                res.Add(new RowData(dd.ToArray()));

            }

            return res;

        }

        public static List<ColData> ToColData(this string[,] data)
        {

            int num = data.GetLength(0);
            List<ColData> res = new List<ColData>();
            for (int j = 0; j < data.GetLength(1); j++)
            {
                List<string> tmp = new List<string>();
                for (int i = 0; i < data.GetLength(0); i++)
                {
                    tmp.Add(data[i, j]);
                }
                res.Add(new ColData(tmp.ToArray()));
            }
            return res;

        }


        public static string[,] ToStringArray(this IEnumerable<RowData> data_)
        {
            List<RowData> data = data_.ToList();
            string[,] res = new string[data.Count, data.Max(t => t.data.Count())];

            for (int i = 0; i < data.Count; i++)
            {
                for (int j = 0; j < data[i].data.Count(); j++)
                {
                    res[i, j] = data[i].data[j];
                }
            }

            return res;
        }


        public static string[,] ToStringArray(this IEnumerable<ColData> data_)
        {
            List<ColData> data = data_.ToList();
            string[,] res = new string[data.Max(t => t.data.Count()), data.Count];
            for (int j = 0; j < data.Count; j++)
            {
                for (int i = 0; i < data[j].data.Count(); i++)
                {
                    res[i, j] = data[j].data[i];
                }
            }
            return res;

        }


    }


    public static class Extension
    {
        public static string[,] DistEmpty(this string[,] data)
        {
            if (null == data) return null;

            int row = data.GetLength(0);
            int col = data.GetLength(1);

            int lastRowNuber = -1;
            for (int i = row - 1; i > 0; i--)
            {
                string tmp = "";
                for (int j = 0; j < data.GetLength(1); j++)
                {
                    tmp = tmp + data[i, j];
                }
                if (tmp.Trim() == "")
                {
                    lastRowNuber = i;
                }
                else
                {
                    break;
                }
            }
            lastRowNuber = lastRowNuber == -1 ? row : lastRowNuber;


            int lastColNuber = -1;
            for (int i = col - 1; i > 0; i--)
            {
                string tmp = "";
                for (int j = 0; j < data.GetLength(0); j++)
                {
                    tmp = tmp + data[j, i];
                }
                if (tmp.Trim() == "")
                {
                    lastColNuber = i;
                }
                else
                {
                    break;
                }
            }
            lastColNuber = lastColNuber == -1 ? col : lastColNuber;

            string[,] res = new string[lastRowNuber, lastColNuber];
            for (int i = 0; i < lastRowNuber; i++)
            {
                for (int j = 0; j < lastColNuber; j++)
                {
                    res[i, j] = data[i, j];
                }
            }

            return res;
        }





        public static object[,] ToObjects(this string[,] data)
        {
            int rowNum = data.GetLength(0);
            int colNum = data.GetLength(1);
            object[,] results = new object[rowNum, colNum];
            for (int i = 0; i < rowNum; i++)
            {
                for (int j = 0; j < colNum; j++)
                {
                    string d = null == data[i, j] ? "" : data[i, j].Trim();
                    bool InNumber = double.TryParse(d, out double num);
                    if (InNumber)
                    {
                        results[i, j] = num;
                    }
                    else
                    {
                        results[i, j] = d;
                    }
                }
            }
            return results;

        }

    }
}
