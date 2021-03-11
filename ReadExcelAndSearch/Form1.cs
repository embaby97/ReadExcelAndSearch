using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ClosedXML.Excel;

namespace ReadExcelAndSearch
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            using(OpenFileDialog ofd =new OpenFileDialog() { Filter= "Excel Workbook|.xls;*.xlsx;*.xlsm", Multiselect = false })
            {//Excel Files|*.xls;*.xlsx;*.xlsm
                //Excel Workbook|*.xlsx
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    Cursor.Current = Cursors.WaitCursor;
                    DataTable dt = new DataTable(); 
                    using (XLWorkbook workbook = new XLWorkbook(ofd.FileName))
                    {
                        bool isFirstRow = true;
                        var rows = workbook.Worksheet(1).RowsUsed();
                        foreach(var row in rows)
                        {
                            if (isFirstRow)
                            {
                                //adding column
                                foreach(IXLCell cell in row.Cells())
                                {
                                    dt.Columns.Add(cell.Value.ToString());
                                }
                                isFirstRow = false;

                            }
                            else
                            {
                                dt.Rows.Add();
                                int i = 0;
                                foreach(IXLCell cell in row.Cells())
                                {
                                    dt.Rows[dt.Rows.Count - 1][i++] = cell.Value.ToString();
                                }
                            }
                        }
                        dataGridView1.DataSource = dt.DefaultView;
                        lblTotal.Text = $"Total records: { dataGridView1.RowCount}";
                        Cursor.Current = Cursors.Default;

                    }
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try 
            {
                DataView dv = dataGridView1.DataSource as DataView;
                if (dv != null)
                    dv.RowFilter =txtFrom.Text ;


            }
            catch (Exception ex){
                MessageBox.Show(ex.Message, "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void txtSearch_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)13)
                btnSearch.PerformClick();
        }

        private void btnSearch_Click(object sender, EventArgs e)
        {
            try
            {
                DataView dv = dataGridView1.DataSource as DataView;

                if (dv != null)
                {
                    // dv.RowFilter = txtSearch.Text;
                    //MessageBox.Show(dataGridView1.Columns[9].HeaderText.ToString());
                    string cul = dataGridView1.Columns[9].HeaderText.ToString();
                    
                    dv.RowFilter = cul + ">= " +txtFrom.Text+ " and " + cul + "<= "+txtTo.Text;

                }

                ///bs.Filter = dataGridView1.Columns[1].HeaderText.ToString() + " LIKE '%" + searchTextBox.Text + "%'";
                /*bs.Filter = string.Format("Date >= #{0:yyyy/MM/dd}# And Date <= #{1:yyyy/MM/dd}#", dateTimePicker1.Value, dateTimePicker2.Value);*/


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
