using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

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

                columnHeaderStyle.BackColor = Color.Beige;
                columnHeaderStyle.Font = new Font("Verdana", 10, FontStyle.Bold);
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
    }
}
