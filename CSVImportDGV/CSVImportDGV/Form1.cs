using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.VisualBasic.FileIO; // This namespace usage is important or else TextFieldParser method will lead to error
using System.IO;

/*
 * Proyecto: Importar/Exportar CSV a DataGridView a Excel
 * Fecha: 25-jul-2021
 * Funcion: Importar archivo de nombre "archivo.csv" a DataGridView y exportarlo a "output.xlsx"
 * Desarrollador: edcruces99
 * 
 * Referencia:
 * 
 * Exportar
 * https://www.c-sharpcorner.com/UploadFile/hrojasara/export-datagridview-to-excel-in-C-Sharp/
 *
 * Importar
 * https://social.technet.microsoft.com/wiki/contents/articles/50922.c-importing-csv-content-and-populating-datagridview-in-windows-form-app-with-the-content.aspx
 * */

namespace CSVImportDGV
{
    public partial class frmImportarCSV : Form
    {
        public frmImportarCSV()
        {
            InitializeComponent();
        }
 
    public class ReadCSV
    {
        public DataTable readCSV;
 
        public ReadCSV(string fileName, bool firstRowContainsFieldNames = true)
        {
            readCSV = GenerateDataTable(fileName, firstRowContainsFieldNames);
        }
 
        private static DataTable GenerateDataTable(string fileName, bool firstRowContainsFieldNames = true)
        {
            DataTable result = new DataTable();
 
            if (fileName == "")
            {
                return result;
            }
 
            string delimiters = ",";
            string extension = Path.GetExtension(fileName);
 
            if (extension.ToLower() == "txt")
                delimiters = "\t";
            else if (extension.ToLower() == "csv")
                delimiters = ",";
 
            using (TextFieldParser tfp = new TextFieldParser(fileName))
            {
                tfp.SetDelimiters(delimiters);
 
                // Get The Column Names
                if (!tfp.EndOfData)
                {
                    string[] fields = tfp.ReadFields();
 
                    for (int i = 0; i < fields.Count(); i++)
                    {
                        if (firstRowContainsFieldNames)
                            result.Columns.Add(fields[i]);
                        else
                            result.Columns.Add("Col" + i);
                    }
 
                    // If first line is data then add it
                    if (!firstRowContainsFieldNames)
                        result.Rows.Add(fields);
                }
 
                // Get Remaining Rows from the CSV
                while (!tfp.EndOfData)
                    result.Rows.Add(tfp.ReadFields());
            }
 
            return result;
        }
    }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog openFileDialog1 = new OpenFileDialog();

                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    lblRuta.Text = openFileDialog1.FileName;

                    ReadCSV csv = new ReadCSV(lblRuta.Text.Trim());

                    try
                    {
                        dataGridView1.DataSource = csv.readCSV;
                    }
                    catch (Exception ex)
                    {
                        throw new Exception(ex.Message);
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        private void btnExportar_Click(object sender, EventArgs e)
        {
            // creating Excel Application  
            Microsoft.Office.Interop.Excel._Application app = new Microsoft.Office.Interop.Excel.Application();
            // creating new WorkBook within Excel application  
            Microsoft.Office.Interop.Excel._Workbook workbook = app.Workbooks.Add(Type.Missing);
            // creating new Excelsheet in workbook  
            Microsoft.Office.Interop.Excel._Worksheet worksheet = null;
            // see the excel sheet behind the program  
            app.Visible = true;
            // get the reference of first sheet. By default its name is Sheet1.  
            // store its reference to worksheet  
            worksheet = workbook.Sheets["Hoja1"];
            worksheet = workbook.ActiveSheet;
            // changing the name of active sheet  
            worksheet.Name = "Exported from gridview";
            // storing header part in Excel  
            for (int i = 1; i < dataGridView1.Columns.Count + 1; i++)
            {
                worksheet.Cells[1, i] = dataGridView1.Columns[i - 1].HeaderText;
            }
            // storing Each row and column value to excel sheet  
            for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
            {
                for (int j = 0; j < dataGridView1.Columns.Count; j++)
                {
                    worksheet.Cells[i + 2, j + 1] = dataGridView1.Rows[i].Cells[j].Value.ToString();
                }
            }
            // save the application  
            workbook.SaveAs(@"D:\Users\eduardo.cruces\Documents\Visual Studio 2013\Projects\CSVImportDGV\CSVImportDGV\bin\Debug\output.xlsx", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            // Exit from the application  
            app.Quit();  
        }
    }
}
