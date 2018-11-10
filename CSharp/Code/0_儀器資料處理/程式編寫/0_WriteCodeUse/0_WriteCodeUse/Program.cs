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
            
            string fileName = "1071013(NewErrorCase).GSI";
            //string fileName = "1061231.GSI";
            string savePath = Path.Combine(oriPath, "AllCase", "4.TheGSI");
            string dataPath = Path.Combine(oriPath, "AllCase", "4.TheGSI");



            string res = "";
            res += TheGSI.TheGSI.TheGSI_Main(savePath, savePath, fileName); 

            ArrayList files = GetFileName_sub(savePath, ".GSI");
            foreach (string ff in files)
            {
                //res += OBDAT.OBDAT.OBMain_sub(oriPath, savePath, ff, "0.009");
                //res += TheCmpExcelData.TheCmpExcelData.TheCmp_Main(savePath, savePath, ff);
                //res += GPS_SORT.GPSSORT.GPSSORT_Main(savePath, savePath, ff); 
                //res += TheZTStoAGA.TheZTStoAGA.TheZTStoAGA_Main(savePath, savePath, ff);
                //res += TheJOBtoT01.TheJOBtoT01.TheJOBtoT01_Main(savePath, savePath, ff); 
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
