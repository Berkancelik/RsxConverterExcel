/* This application reads .Net resource files from specified folder and then write its contents to an excel file.
 * If you want you can to generate one excel file for one resource file or you can generate a single excel file 
 * and one sheet for one resource file whithin it.
 */

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Excel;
using System.Runtime.InteropServices;
using System.Reflection;
using System.IO;
using System.Collections;

namespace ResourceToExcelConveretr
{
    public partial class Form1 : Form
    {
        // reference to Microsoft Excel object
        Excel.Application exlObj = null;
        
        public Form1()
        {
            InitializeComponent();
            progressBar1.Visible = false;
            label3.Visible = false;
        }

        private void btnSourceBrowse_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                txtSourcePath.Text = folderBrowserDialog1.SelectedPath;
            }
        }

        private void btnTargetBrowse_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                txtTargetPath.Text = folderBrowserDialog1.SelectedPath;
            }
        }

        private void btnStart_Click(object sender, EventArgs e)
        {
            string fileName = string.Empty;
            int noOfFilesWritten = 0;

            // initializing the excel application
            exlObj = new Excel.Application();
            // Showing the error message if any problem occures in starting excel application
            if (exlObj == null)
            {
                MessageBox.Show("Problem in starting Excel.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                System.Windows.Forms.Application.Exit();
            }
            // you can keep the excel application visible or invisible while writing to it
            exlObj.Visible = false;
                        
            if (txtSourcePath.Text == "" || txtTargetPath.Text == "")
            {
                MessageBox.Show("Please enter the paths.");
                return;
            }

            btnSourceBrowse.Enabled = false;
            btnTargetBrowse.Enabled = false;
            btnStart.Enabled = false;
            // getting .resx file collection from given directory
            string []fileList = Directory.GetFiles(txtSourcePath.Text,"*.resx");
                        
            label3.Visible = true;
            progressBar1.Maximum = fileList.Length;
            progressBar1.Visible = true;
            this.Refresh();

            ////
            // creating excel workbook
            Excel.Workbooks workbooks = exlObj.Workbooks;
            // "workbook.Add" method returns a workbook with specifide template
            _Workbook workbook = workbooks.Add(XlWBATemplate.xlWBATWorksheet);
            // getting the collection of excel sheets from workbook however right now there is only one worksheet in workbook
            Sheets sheets = workbook.Worksheets;
            ////

            // iterating file collection
            foreach (string filePath in fileList)
            {
                fileName = Path.GetFileName(filePath);
                // creating RESX resource reader that will read the .resx file
                System.Resources.ResXResourceReader reader = new System.Resources.ResXResourceReader(filePath);
                // creating enumerator to read Resource file line by line
                IDictionaryEnumerator dictionary = reader.GetEnumerator();
               
                GenerateExcelFile(sheets, dictionary, fileName);
                noOfFilesWritten++;
            }
            // if only sigle excel file is generated than we need to save it only once
            if (chkGenerateSingle.Checked)
            {
                workbook.SaveAs(txtTargetPath.Text + "\\Excel_Resources.xls", 1, null, null, false, false,XlSaveAsAccessMode.xlNoChange, null, null, null,null,null);
            }           
            
            btnSourceBrowse.Enabled = true;
            btnTargetBrowse.Enabled = true;
            btnStart.Enabled = true;
            // closing the excel application it is necessary otherwise Excel applicatio will continue to run
            exlObj.Application.Quit();
            exlObj = null;
            MessageBox.Show("Operation successfully completed.\nThe total No. of files written is " + noOfFilesWritten +".");
            label3.Visible = false;            
            progressBar1.Visible = false;
            progressBar1.Value = 0;
        }


        public void GenerateExcelFile(Sheets sheets, IDictionaryEnumerator dictionary, string fileName)
        {
            // If single excel file is to be generated than we need to add new sheet every time a new resorce file is written
            if (chkGenerateSingle.Checked)
            {
                // this method adds a new sheet to Sheet collection of workbook
                // here Missing.Value signifies that these arguments are unknown
                // we can also pass appropriate arguments if we wish
                sheets.Add(Missing.Value, Missing.Value, Missing.Value, Missing.Value);                                         
            }

            // getting the worksheet from Sheet's collection. keep in mind in Excel index starts from 1
            // and index 1 always gives the newly added worksheet
            _Worksheet worksheet = (_Worksheet)sheets.get_Item(1);               
            
            // naming the worksheet. the name of worksheet should not exceed 31 characters
            if (fileName.Length > 30)
            {
                worksheet.Name = fileName.Substring(0, 29);
            }
            else
            {
                worksheet.Name = fileName;
            }

            // get_range gives the object of a particular row from index1 to index2
            Range range = worksheet.get_Range("A1", "B1");

            // applying text format on a perticular range. these can be applied on hwole worksheet too
            range.Font.Bold = true;
            // writing text to a cell
            range.set_Item(1, 1, "KEY");
            range.set_Item(1, 2, "VALUE");
            // applying format on worksheet
            worksheet.Columns.ColumnWidth = 40;
            worksheet.Rows.WrapText = true;

            int count = 2;

            // enumerating item collection in resource file and writing to worksheet.
            while (dictionary.MoveNext())
            {
                // if we wish we can avoid writing empty values or can keep writing
                //if (dictionary.Value.ToString() == "")
                //{
                //    continue;
                //}
                range.set_Item(count, 1, dictionary.Key);
                range.set_Item(count, 2, dictionary.Value);
                count++;
            }
            // If we are generating different excel files for each RESX file we need to save it every time
            if (!chkGenerateSingle.Checked)
            {
                worksheet.SaveAs(txtTargetPath.Text + "\\" + fileName + ".xls", 1, null, null, false, false, null, null, null, null);                
            }
            progressBar1.Value += 1;
        }
    }
}