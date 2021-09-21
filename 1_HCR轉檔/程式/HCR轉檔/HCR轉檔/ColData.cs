using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HCR轉檔
{
    [DebuggerDisplay("Count = {data.Length}, head = {head}")]
    public class ColData
    {
        public string head = "";
        public string[] data;
        public ColData(string[] data)
        {
            this.data = data;
            
        }

    }
}
