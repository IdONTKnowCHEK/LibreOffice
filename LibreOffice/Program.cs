using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace LibreOffice
{
    internal static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            string param = args.Length > 0 ? args[0] : string.Empty;
            // Application.Run(new Form1(param));
            LOCalc.Class1 loCalc = new LOCalc.Class1();
            if (loCalc.bOpenOfficeInstalled == false)
                return;
            loCalc.bStartOpenOfficeLoader();
            loCalc.bCreateWorkbook();

            if (param != string.Empty)
            {
                loCalc.bSetText("A1", param);
            }
            else
            {
                loCalc.bSetValue("A1", 12345555);
                loCalc.bSetValue("B1", 2468);
                loCalc.bSetFormula("C1", "=A1+B1");
                loCalc.bSetText("A2", "Text entered here.");
                loCalc.bSetDate("A3", 2023, 12, 25);
                loCalc.bSetBackgroundColor("A3", Color.Red);
            }

            string filePath = "C:\\file\\file.ods";
            string strFilePath = loCalc.strSaveWorkbook(filePath);
            Console.WriteLine("Saved as: " + strFilePath);
            loCalc.bCloseWorkbook();

            Application.Exit();
        }
    }
}
