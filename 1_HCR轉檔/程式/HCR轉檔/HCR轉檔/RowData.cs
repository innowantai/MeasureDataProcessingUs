using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HCR轉檔
{
    [DebuggerDisplay("Count = {data.Length}")]
    public class RowData
    {

        public string[] data;
        public RowData(string[] data)
        {
            this.data = data;
        }

        
    }
}
