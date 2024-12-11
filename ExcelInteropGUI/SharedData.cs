using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelInteropGUI
{
    public static class SharedData
    {
        public static List<(int ChangedRow, int ChangedCol, object ChangedVal)> Log { get; set; } = new List<(int ChangedRow, int ChangedCol, object ChangedVal)>();
        public static List<(object ChangedVal, DateTime EdittedTime)> NewValues { get; set; } = new List<(object ChangedVal, DateTime EdittedTime)>();

    }

    public static class ChangeLanguage
    {
        public static void change(string lang, Control control)
        {
            ComponentResourceManager componentResourceManager = new ComponentResourceManager(typeof(Languages));
            foreach (System.Windows.Forms.Control c in control.Controls)
            {
                componentResourceManager.ApplyResources(c, c.Name, new System.Globalization.CultureInfo(lang));
            }
        }
    }
}
