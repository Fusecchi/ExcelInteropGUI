using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelInteropGUI
{
    public static class SharedData
    {
        public static List<(int ChangedRow, int ChangedCol, object ChangedVal)> Log { get; set; } = new List<(int ChangedRow, int ChangedCol, object ChangedVal)>();
        public static List<(object ChangedVal, DateTime EdittedTime)> NewValues { get; set; } = new List<(object ChangedVal, DateTime EdittedTime)>();
        
    }
}
