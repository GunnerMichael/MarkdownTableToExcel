using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using Excel = Microsoft.Office.Interop.Excel;

namespace MarkdownTableToExcel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void toExcelCommand_Click(object sender, EventArgs e)
        {
            string text = markdownBox.Text;

            string[] lines = text.Split("\r\n");

            bool foundHeader = false;

            List<string> columns = new List<string>();

            List<List<string>> data = new List<List<string>>();

            foreach (var line in lines)
            {
                if (line.StartsWith("|-"))
                {
                    continue;
                }

                if (line.StartsWith("|"))
                {
                    if (foundHeader == false)
                    {
                        string[] cols = line.Substring(line.IndexOf("|") + 1).Split("|");

                        foreach (var item in cols)
                        {
                            if (string.IsNullOrEmpty(item) == false)
                            {
                                columns.Add(item);
                            }
                        }

                        foundHeader = true;

                    }
                    else
                    {
                        string[] cols = line.Substring(line.IndexOf("|") + 1).Split("|");

                        List<string> row = new List<string>();

                        foreach (var item in cols)
                        {
                            var itemText = item.Trim();

                            if (string.IsNullOrEmpty(itemText) == false && itemText.Contains("----") == false)
                            {
                                row.Add(itemText);
                            }
                        }

                        data.Add(row);
                    }
                }
            }




            var excelApp = new Excel.Application();
            //Make the object visible.
            excelApp.Visible = true;

            // Create a new, empty workbook and add it to the collection returned
            // by property Workbooks. The new workbook becomes the active workbook.
            // Add has an optional parameter for specifying a praticular template.
            // Because no argument is sent in this example, Add creates a new workbook.
            excelApp.Workbooks.Add();

            // This example uses a single workSheet. The explicit type casting is
            // removed in a later procedure.
            Excel._Worksheet workSheet = (Excel.Worksheet)excelApp.ActiveSheet;

            char cell = 'A';

            foreach (var item in columns)
            {
                workSheet.Cells[1, $"{cell}"] = item;
                cell++;
            }

            //workSheet.Cells[1, "A"] = "ID Number";
            //workSheet.Cells[1, "B"] = "Current Balance";

            var rowCount = 2;

            foreach(var rowData in data)
            {
                cell = 'A';
                foreach (var col in rowData)
                {
                    workSheet.Cells[rowCount, $"{cell}"] = col;
                    cell++;
                }

                rowCount++;
            }

            //foreach (var acct in accounts)
            //{
            //    row++;
            //    workSheet.Cells[row, "A"] = acct.ID;
            //    workSheet.Cells[row, "B"] = acct.Balance;
            //}

            for (int i = 1; i <= columns.Count; i++)
            {
                workSheet.Columns[i].AutoFit();
            }
            //workSheet.Columns[2].AutoFit();

            //// Call to AutoFormat in Visual C#. This statement replaces the
            //// two calls to AutoFit.
            //workSheet.Range["A1", "B3"].AutoFormat(
            //    Excel.XlRangeAutoFormat.xlRangeAutoFormatClassic2);

            //// Put the spreadsheet contents on the clipboard. The Copy method has one
            //// optional parameter for specifying a destination. Because no argument
            //// is sent, the destination is the Clipboard.
            //workSheet.Range["A1:B3"].Copy();
        }
    }
}
