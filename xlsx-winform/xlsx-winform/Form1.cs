using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using DocumentFormat.OpenXml.Bibliography;
using Excel = ClosedXML.Excel;

namespace xlsx_winform
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            // read from file
            string CurrentDir = System.IO.Path.GetDirectoryName(Assembly.GetEntryAssembly().Location);
            CurrentDir += @"\sheet.xlsx";
            var wb = new Excel.XLWorkbook(CurrentDir);
            var ws = wb.Worksheets.Worksheet("Sheet1");
            if (System.IO.File.Exists(CurrentDir)) {
                TextBox[] FormA = new[] { SprA1, SprA2, SprA3, SprA4 };
                TextBox[] FormB = new[] { SprB1, SprB2, SprB3, SprB4 };
                TextBox[] FormC = new[] { SprC1, SprC2, SprC3, SprC4 };
                TextBox[] FormD = new[] { SprD1, SprD2, SprD3, SprD4 };
                Object[] SheetA = new[] { ws.Cell("A1").Value, ws.Cell("A2").Value, ws.Cell("A3").Value, ws.Cell("A4").Value };
                Object[] SheetB = new[] { ws.Cell("B1").Value, ws.Cell("B2").Value, ws.Cell("B3").Value, ws.Cell("B4").Value };
                Object[] SheetC = new[] { ws.Cell("C1").Value, ws.Cell("C2").Value, ws.Cell("C3").Value, ws.Cell("C4").Value };
                Object[] SheetD = new[] { ws.Cell("D1").Value, ws.Cell("D2").Value, ws.Cell("D3").Value, ws.Cell("D4").Value };
                var SpreadA = FormA.Zip(SheetA, (f,s) => new { FrmA = f, ShtA = s });
                foreach (var fs in SpreadA)
                {
                    fs.FrmA.Text = Convert.ToString(fs.ShtA);
                }
                var SpreadB = FormB.Zip(SheetB, (f, s) => new { FrmB = f, ShtB = s });
                foreach (var fs in SpreadB)
                {
                    fs.FrmB.Text = Convert.ToString(fs.ShtB);
                }
                var SpreadC = FormC.Zip(SheetC, (f, s) => new { FrmC = f, ShtC = s });
                foreach (var fs in SpreadC)
                {
                    fs.FrmC.Text = Convert.ToString(fs.ShtC);
                }
                var SpreadD = FormD.Zip(SheetD, (f, s) => new { FrmD = f, ShtD = s });
                foreach (var fs in SpreadD)
                {
                    fs.FrmD.Text = Convert.ToString(fs.ShtD);
                }
            } else
            {
                MessageBox.Show("File doesn't exist.","Error!",MessageBoxButtons.OK,MessageBoxIcon.Error);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // write to file
            string CurrentDir = System.IO.Path.GetDirectoryName(Assembly.GetEntryAssembly().Location);
            CurrentDir += @"\sheet.xlsx";
            var wb = new Excel.XLWorkbook(CurrentDir);
            var ws = wb.Worksheets.Worksheet("Sheet1");
            if (System.IO.File.Exists(CurrentDir))
            {
                TextBox[] FormA = new[] { SprA1, SprA2, SprA3, SprA4 };
                TextBox[] FormB = new[] { SprB1, SprB2, SprB3, SprB4 };
                TextBox[] FormC = new[] { SprC1, SprC2, SprC3, SprC4 };
                TextBox[] FormD = new[] { SprD1, SprD2, SprD3, SprD4 };
                String[] SheetA = new[] { "A1", "A2", "A3", "A4" };
                String[] SheetB = new[] { "B1", "B2", "B3", "B4" };
                String[] SheetC = new[] { "C1", "C2", "C3", "C4" };
                String[] SheetD = new[] { "D1", "D2", "D3", "D4" };

                // what the hell is this
                for (int i = 0; i < FormA.Length; i++)
                {
                    ws.Cell(SheetA[i]).Value = FormA[i].Text;
                }
                for (int i = 0; i < FormB.Length; i++)
                {
                    ws.Cell(SheetB[i]).Value = FormB[i].Text;
                }
                for (int i = 0; i < FormC.Length; i++)
                {
                    ws.Cell(SheetC[i]).Value = FormC[i].Text;
                }
                for (int i = 0; i < FormD.Length; i++)
                {
                    ws.Cell(SheetD[i]).Value = FormD[i].Text;
                }

                wb.Save();
            }
            else
            {
                MessageBox.Show("File doesn't exist.", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
