using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace WindowsFormsApplication1
{
    public partial class FormWindow : Form
    {
        // the list used for saving there the only workbook to work on it with another private function

        List<Excel.Workbook> booksList = new List<Excel.Workbook>();

        public FormWindow()
        {
            InitializeComponent();
            textBox1.Text = "Wpisz indeks";
        }

        private void button1_Click(object sender, System.EventArgs e)
        {
            booksList.Clear();
            Excel.Application oXL;
            Excel.Workbook oWB;
            try
            {
                // opening the file, not only getting the path but visually opening the Excel application
                // and working on the open application on live

                string fileName = "";
                if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    string path = openFileDialog1.FileName;
                    fileName = path as string;
                    textBox2.Text = fileName;
                }
                oXL = new Excel.Application();
                oXL.Visible = true;

                oWB = (Excel.Workbook)(oXL.Workbooks.Open(fileName,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing));

                oXL.Visible = true;
                oXL.UserControl = true;
                booksList.Add(oWB);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            int checkIfInTheWorkBook = 0;

            try
            {
                // getting the workbook, that is already on the list and looking for typed value within it

                Excel.Workbook oWB;
                oWB = booksList[0];
                Excel.Worksheet oSheet;
                int sheetsNumber = oWB.Worksheets.Count;

                for (int i = 1; i < sheetsNumber + 1; i++)
                {
                    oSheet = (Excel.Worksheet)oWB.Worksheets[i];
                    Excel.Range usedRange = oSheet.UsedRange;
                    Excel.Range rows = usedRange.Rows;
                    Excel.Range cols = usedRange.Columns;

                    foreach (Excel.Range row in rows)
                    {
                        foreach (Excel.Range cell in row)
                        {
                            string cellVal1 = cell.Value.ToString();
                            string cellVal2 = cell.Value2.ToString();
                            if ((cellVal2 == textBox1.Text.ToString() || cellVal1 == textBox1.Text.ToString()) && textBox1.Text != "")
                            {
                                cell.Interior.Color = System.Drawing.Color.Black;
                                textBox1.Text = "Wpisz indeks";
                                checkIfInTheWorkBook += 1;
                            }
                        }
                    }
                }
                if (checkIfInTheWorkBook == 0)
                {
                    MessageBox.Show("Indeks nie został znaleziony!");
                    textBox1.Text = "Wpisz indeks";
                }
                else if(checkIfInTheWorkBook > 0)
                {
                    MessageBox.Show("Odznaczono wybrany indeks x"+checkIfInTheWorkBook);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
    }
}
