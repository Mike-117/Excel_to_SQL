// add NuGet ExcelDataReader and ExcelDataReader.DataSet
using ExcelDataReader;
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

namespace ExcelToSQLWinForm
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        // select the table that contains the desired data, and bind it to the DataGridView
        private void cboSheet_SelectedIndexChanged(object sender, EventArgs e)
        {
            DataTable myTable = tableCollection[cboSheet.SelectedItem.ToString()];
            dataGridView1.DataSource = myTable;
        }

        DataTableCollection tableCollection;

        //click event handler allows user to choose the Excel file
        private void btnBrowse_Click(object sender, EventArgs e)
        {
            // opens the window to browse through and choose file
            using (OpenFileDialog openFileDialog = new OpenFileDialog() { Filter = "Excel Document|*.xls; *.xlsx" })
            {
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    textFilename.Text = openFileDialog.FileName;
                    using (var stream = File.Open(openFileDialog.FileName, FileMode.Open, FileAccess.Read))
                    {
                        // CreateReader method allows user to read the excel file, the data is returned as a DataSet
                        using (IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream))
                        {
                            DataSet result = reader.AsDataSet(new ExcelDataSetConfiguration()
                            {
                                ConfigureDataTable = (_) => new ExcelDataTableConfiguration() { UseHeaderRow = true }
                            });
                            // read all the table names in the DataSet
                            tableCollection = result.Tables;
                            cboSheet.Items.Clear();
                            // add table to the combobox
                            foreach (DataTable myTable in tableCollection)
                            {
                                cboSheet.Items.Add(myTable.TableName);
                            }
                        }
                    }

                }
            }
        }
    }
}
