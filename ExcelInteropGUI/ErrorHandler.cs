using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelInteropGUI
{
    internal static class ErrorHandling
    {
        public static void Initialize()
        {
            Application.ThreadException += (sender, e) =>
            {
                MessageBox.Show($"An Error Occured: {e.Exception.Message}");
            };

            AppDomain.CurrentDomain.UnhandledException += (sender, e) => 
            {
                Exception ex = (Exception)e.ExceptionObject;
                MessageBox.Show($"An Error Occured: {ex.Message}");
            };
        }
    }
}
