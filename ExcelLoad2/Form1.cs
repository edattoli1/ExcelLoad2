using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelLoad2
{
    public partial class Form1 : Form
    {

        public Form1()
        {
            InitializeComponent();
            openFileDialog1.InitialDirectory = System.IO.Path.GetDirectoryName(Application.ExecutablePath);
        }

        //public List<List<double>> table;
        public List<List<string>> table;

        //private DataGridView dataGridView1 = new DataGridView();
        //private BindingSource bindingSource1 = new BindingSource();

        private void button1_Click(object sender, EventArgs e)
        {
            DialogResult result = openFileDialog1.ShowDialog();
            if (result == DialogResult.OK) // Test result.
            {
                //List<double> row1 = new List<double> {1,2,3};
                //List<double> row2 = new List<double> { 1, 2, 3 };

                //List<string> row1 = new List<string> { "1", "2", "3" };
               //List<string> row2 = new List<string> { "3", "4", "5" };

               // double [] row1 = {1,2,3};
               // double [] row2 = {4,5,6};

                //table= new List<List<String>> ();
            
                //table.Add(row1);
                //table.Add(row2);
                //table.Add(row1);

                //Make DataGrid Uneditable
                //dataGridView1.Enabled = false;
                dataGridView1.ReadOnly = true;

                // Create an unbound DataGridView by declaring a column count.
                dataGridView1.ColumnCount = 3;
                dataGridView1.ColumnHeadersVisible = true;

                // Set the column header style.
                DataGridViewCellStyle columnHeaderStyle = new DataGridViewCellStyle();

                columnHeaderStyle.BackColor = System.Drawing.Color.Beige;
                columnHeaderStyle.Font = new System.Drawing.Font("Verdana", 10, FontStyle.Bold);
                dataGridView1.ColumnHeadersDefaultCellStyle = columnHeaderStyle;

                // Set the column header names.
                dataGridView1.Columns[0].Name = "Time (min)";
                dataGridView1.Columns[1].Name = "MFC1";
                dataGridView1.Columns[2].Name = "MFC2";

                // Populate the rows. 
            //    string[] row1 = new string[] { "Meatloaf", "Main Dish", "ground beef",
            //"**" };
            //    string[] row2 = new string[] { "Key Lime Pie", "Dessert", 
            //"lime juice, evaporated milk", "****" };
            //    string[] row3 = new string[] { "Orange-Salsa Pork Chops", "Main Dish", 
            //"pork chops, salsa, orange juice", "****" };
            //    string[] row4 = new string[] { "Black Bean and Rice Salad", "Salad", 
            //"black beans, brown rice", "****" };
            //    string[] row5 = new string[] { "Chocolate Cheesecake", "Dessert", 
            //"cream cheese", "***" };
            //    string[] row6 = new string[] { "Black Bean Dip", "Appetizer", 
            //"black beans, sour cream", "***" };
            //    object[] rows = new object[] { row1, row2, row3, row4, row5, row6 };

                string[] row1 =  { "1", "2", "3" };
                string[] row2 = { "3", "4", "5" };

                object[] rows = new object[] { row1, row2 };

                foreach (string[] rowArray in rows)
                {
                    dataGridView1.Rows.Add(rowArray);
                }
            }
        }

        public static string XLGetCellValue(string fileName, string sheetName, string addressName)
        {
            string value = null;

            using (SpreadsheetDocument document = SpreadsheetDocument.Open(fileName, false))
            {
                WorkbookPart wbPart = document.WorkbookPart;

                // Find the sheet with the supplied name, and then use that Sheet object
                // to retrieve a reference to the appropriate worksheet.
                Sheet theSheet = wbPart.Workbook.Descendants<Sheet>().
                  Where(s => s.Name == sheetName).FirstOrDefault();

                if (theSheet == null)
                {
                    throw new ArgumentException("sheetName");
                }

                // Retrieve a reference to the worksheet part, and then use its Worksheet property to get 
                // a reference to the cell whose address matches the address you've supplied:
                WorksheetPart wsPart = (WorksheetPart)(wbPart.GetPartById(theSheet.Id));
                Cell theCell = wsPart.Worksheet.Descendants<Cell>().
                  Where(c => c.CellReference == addressName).FirstOrDefault();

                // If the cell doesn't exist, return an empty string:
                if (theCell != null)
                {
                    //value = theCell.InnerText;

                    /*  To retrieve just the value */
                    value = theCell.CellValue.InnerText;


                    // If the cell represents an integer number, you're done. 
                    // For dates, this code returns the serialized value that 
                    // represents the date. The code handles strings and booleans
                    // individually. For shared strings, the code looks up the corresponding
                    // value in the shared string table. For booleans, the code converts 
                    // the value into t he words TRUE or FALSE.
                    if (theCell.DataType != null)
                    {
                        switch (theCell.DataType.Value)
                        {
                            case CellValues.SharedString:
                                // For shared strings, look up the value in the shared strings table.
                                var stringTable = wbPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                                // If the shared string table is missing, something's wrong.
                                // Just return the index that you found in the cell.
                                // Otherwise, look up the correct text in the table.
                                if (stringTable != null)
                                {
                                    value = stringTable.SharedStringTable.ElementAt(int.Parse(value)).InnerText;
                                }
                                break;

                            case CellValues.Boolean:
                                switch (value)
                                {
                                    case "0":
                                        value = "FALSE";
                                        break;
                                    default:
                                        value = "TRUE";
                                        break;
                                }
                                break;
                        }
                    }
                }
            }
            return value;
        }

    }
}
