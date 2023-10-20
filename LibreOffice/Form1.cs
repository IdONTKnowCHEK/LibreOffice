using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace LibreOffice
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            LOCalc.Class1 loCalc = new LOCalc.Class1();
            if (loCalc.bOpenOfficeInstalled == false)
                return;
            loCalc.bStartOpenOfficeLoader();
            loCalc.bCreateWorkbook();
            loCalc.bSetValue("A1", 12345555);
            loCalc.bSetValue("B1", 2468);
            loCalc.bSetFormula("C1", "=A1+B1");
            loCalc.bSetText("A2", "Text entered here.");
            loCalc.bSetDate("A3", 2023, 12, 25);
            loCalc.bSetBackgroundColor("A3", Color.Red);
            string filePath = "C:\\file\\file.ods";
            string strFilePath = loCalc.strSaveWorkbook(filePath);
            Console.WriteLine("Saved as: " + strFilePath);

            loCalc.bCloseWorkbook();
        }
    }
}
