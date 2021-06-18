using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Windows.Forms;
using ExcelDataReader;
using LiveCharts;
using LiveCharts.Wpf;


namespace Table
{
    public partial class Form1 : Form
    {
        private string fileName = string.Empty;
        private DataTableCollection tableCollection = null;

        int indexStudent =0;

        DataSet db;
        public Form1()
        {
            InitializeComponent();
        }
         private void button1_Click(object sender, EventArgs e)
         {
   
               DialogResult rs = openFileDialog1.ShowDialog();

               if (rs == DialogResult.OK)
               {
                  fileName = openFileDialog1.FileName;

                openExelTable(fileName);
               }

         }
        private void openExelTable(string path)
        {
            FileStream stream = File.Open(path, FileMode.Open, FileAccess.Read);

            IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream);

            db = reader.AsDataSet(new ExcelDataSetConfiguration()
            {
                ConfigureDataTable = (x) => new ExcelDataTableConfiguration()
                {
                    UseHeaderRow = true
                }
            });

            tableCollection = db.Tables;
            toolStripComboBox1.Items.Clear();
            foreach(DataTable tabe in tableCollection)
            {
                toolStripComboBox1.Items.Add(tabe.TableName);
            }
            toolStripComboBox1.SelectedIndex = 0;
        }

        private void toolStripComboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            DataTable table = tableCollection[Convert.ToString(toolStripComboBox1.SelectedItem)];
            dataGridView1.DataSource = table;

        }

      

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
        }

        private void button3_Click(object sender, EventArgs e)
        {

            int indexStudent = Convert.ToInt32(textBox1.Text)-1;

            double[] ozenki = new double[dataGridView1.ColumnCount-3];
            int a = 0;

            chart1.Series[0].Points.Clear();

            for (int col = 2; col < dataGridView1.Rows[indexStudent].Cells.Count-1; col++)
            {
                string value = dataGridView1.Rows[indexStudent].Cells[col].Value.ToString();
                ozenki[a] = Convert.ToDouble(value);
                a++;
            }

            for(int i=0;i< dataGridView1.ColumnCount - 3; i++)
            {
                chart1.Series[0].Points.AddXY(i, ozenki[i]);
            }
            

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

            switch(comboBox1.SelectedIndex)
            {
                case 0:
                    chart1.Series[0].Points.Clear();
                    double[] itog = new double[dataGridView1.RowCount];

                    for (int rows=0 ; rows < dataGridView1.Rows.Count-1; rows++)
                    {
                        string items = dataGridView1.Rows[rows].Cells[dataGridView1.ColumnCount-1].Value.ToString();
                        itog[rows] = Convert.ToDouble(items);
                        chart1.Series[0].Points.AddXY(rows, itog[rows]);
                    }


                    break;

                case 1:
                    chart1.Series[0].Points.Clear();
                    double[] iitog = new double[dataGridView1.RowCount];
                    int a = 0;
                    for (int i = 0; i < dataGridView1.RowCount-1; i++)
                    {
                        string value = dataGridView1.Rows[i].Cells[0].Value.ToString();
                        if (value == "ПРИ-311")
                        {
                            string items = dataGridView1.Rows[i].Cells[dataGridView1.ColumnCount-1].Value.ToString();
                            iitog[a] = Convert.ToDouble(items);
                            a++;
                        }
                    }

                    for (int i = 0; i <a; i++)
                    {
                        chart1.Series[0].Points.AddXY(i, iitog[i]);
                    }

                    break;

                case 2:
                    chart1.Series[0].Points.Clear();
                    double[] atog = new double[dataGridView1.RowCount];
                    int b = 0;
                    for(int i =0; i<dataGridView1.RowCount-1;i++)
                    {
                        string value = dataGridView1.Rows[i].Cells[0].Value.ToString();
                        if(value == "ПРИ-312")
                        {
                            string items = dataGridView1.Rows[i].Cells[dataGridView1.ColumnCount-1].Value.ToString();
                            atog[b] = Convert.ToDouble(items);
                            b++;
                        }
                    }
                    for (int i = 0; i <b; i++)
                    {
                        chart1.Series[0].Points.AddXY(i, atog[i]);
                    }
                    break;

            }
        }
    }
}
