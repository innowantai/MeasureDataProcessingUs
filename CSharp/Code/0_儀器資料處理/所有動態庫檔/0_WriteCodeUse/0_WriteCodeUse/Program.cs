using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Collections;
// test candelete
namespace _0_WriteCodeUse
{
    public class Program
    {
        public static string oriPath = System.Environment.CurrentDirectory;
        static void Main(string[] args)
        {
            
           // string savePath = Path.Combine(oriPath, "AllCase", "0.DAT");
            string savePath = Path.Combine(oriPath, "AllCase", "8.TheJOBtoT01");
            string dataPath = savePath;




            ArrayList files = GetFileName_sub(savePath, ".JOB");
            string res = ""; 
            foreach (string ff in files)
            {

                //res += OBDAT.OBDAT.OBMain_sub(oriPath, savePath, ff, "0.009");
                //res += TheCmpExcelData.TheCmpExcelData.TheCmp_Main(savePath, savePath, ff);
                //res += GPS_SORT.GPSSORT.GPSSORT_Main(savePath, savePath, ff); 
                //res += TheZTStoAGA.TheZTStoAGA.TheZTStoAGA_Main(savePath, savePath, ff);
                res += TheJOBtoT01.TheJOBtoT01.TheJOBtoT01_Main(savePath, savePath, ff); 
            }

            Console.WriteLine(res);


            

        }


        /// <summary>
        /// Get files from indicated Path
        /// </summary>
        /// <param name="Path"></param>
        /// <returns></returns>
        public static ArrayList GetFileName_sub(string Path, string fileSubName)
        {
            DirectoryInfo Dir = new DirectoryInfo(Path);
            ArrayList FIleName = new ArrayList();

            foreach (FileInfo f in Dir.GetFiles()) //查詢附檔名為""的文件  
            {
                string index = f.ToString();
                if (index.Contains(fileSubName.ToLower()) | index.Contains(fileSubName.ToUpper()))
                {
                    FIleName.Add(index);
                }
            }
            return FIleName;
        }
    }
}
