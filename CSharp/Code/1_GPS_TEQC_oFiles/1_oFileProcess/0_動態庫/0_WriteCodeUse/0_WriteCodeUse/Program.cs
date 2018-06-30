using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace _0_WriteCodeUse
{
    class Program
    {
        public static string oriPath = System.Environment.CurrentDirectory;

        static void Main(string[] args)
        {
            string res = "";
            string dataPath = Path.Combine(oriPath, "07");
            string fileName = Path.Combine(dataPath, "儀高化算表C_sort_1.xls");
            res = GPSoFileProcess.GPSoFiles.GPSoFile_Main(fileName,dataPath,oriPath,dataPath);

        }
    }
}
