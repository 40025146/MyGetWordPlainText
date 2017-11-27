/****************************** Module Header ******************************\
* Module Name:  MainForm.cs
* Project:      CSOpenXmlGetPlainText
* Copyright(c)  Microsoft Corporation.
* 
* This is the main form of this application. It is used to initialize the UI and 
* handle the events.
*
* This source is subject to the Microsoft Public License.
* See http://www.microsoft.com/en-us/openness/licenses.aspx.
* All other rights reserved.
* 
* THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, 
* EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED 
* WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
\***************************************************************************/



using System;
using System.Windows.Forms;
using System.IO;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ClosedXML.Excel;
using System.Text.RegularExpressions;

namespace CSUsingOpenXmlPlainText
{
    public partial class MainForm : Form
    {
        GetWordPlainText getWordPlainText = null;
        XLWorkbook workbook = null;
        IXLWorksheet worksheet = null;
        static string regex_string = "";
        static string regex_condition = @"[\W_]+";
        public MainForm()
        {
            InitializeComponent();
            this.btnSaveas.Enabled = false;
        }

        /// <summary>
        ///  Handle the btnOpen Click event to load an Word file.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnOpen_Click(object sender, EventArgs e)
        {
            SelectWordFile(); 
        }

        /// <summary>
        /// Show an OpenFileDialog to select a Word document.
        /// </summary>
        /// <returns>
        /// The file name.
        /// </returns>
        private string SelectWordFile()
        {
            string fileName = null;
            using (OpenFileDialog dialog = new OpenFileDialog())
            {
                dialog.Filter = "Word document (*.docx)|*.docx";
                dialog.InitialDirectory = Environment.CurrentDirectory;

                // Retore the directory before closing
                dialog.RestoreDirectory = true;
                if (dialog.ShowDialog()== DialogResult.OK)
                {
                    fileName = dialog.FileName;
                    tbxFile.Text = dialog.FileName;
                    rtbText.Clear();
                }
            }

            return fileName;
        }

        /// <summary>
        /// Get Plain Text from Word file
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnGetPlainText_Click(object sender, EventArgs e)
        {
            try
            {
                getWordPlainText = new GetWordPlainText(tbxFile.Text);
                this.rtbText.Clear();
                this.rtbText.Text = getWordPlainText.ReadWordDocument(txtExcelPath.Text);

                // After read text in word document successfully，make "save as text" button to be enabled.
                this.btnSaveas.Enabled = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            finally
            {
                if (getWordPlainText != null)
                {
                    getWordPlainText.Dispose();
                }
            }
        }

        /// <summary>
        ///  Save the text to text file
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnSaveas_Click(object sender, EventArgs e)
        {
            using (SaveFileDialog savefileDialog = new SaveFileDialog())
            {
                savefileDialog.Filter = "txt document(*.txt)|*.txt";

                // default file name extension
                savefileDialog.DefaultExt = ".txt";

                // Retore the directory before closing
                savefileDialog.RestoreDirectory = true;
                if (savefileDialog.ShowDialog() == DialogResult.OK)
                {
                    string filename = savefileDialog.FileName;
                    rtbText.SaveFile(filename, RichTextBoxStreamType.PlainText);
                    MessageBox.Show("Save Text file successfully, the file path is： " + filename);
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            
            workbook = new XLWorkbook(txtExcelPath.Text);
            worksheet = workbook.Worksheet(1);
            int lastRow = worksheet.LastRow().RowNumber();
            RemoveSpace(2, 1000);
            Reset(2, 1000);
            RemoveSpace(3, 1000);
            Reset(3, 1000);
            Reset(4, 1000);
            worksheet.Columns().Width = 10;
            workbook.Save();
            MessageBox.Show("結束");
        }
        private void RemoveSpace(int col_num,int maxRow)
        {
            for (int i = 2; i < maxRow; i++)
            {
                string str = worksheet.Cell(i, col_num).Value.ToString().Trim();
                worksheet.Cell(i, col_num).Value = str;
                str = str.Replace(" ","");
                worksheet.Cell(i, col_num).Value = str;
            }
        }
    
        private void Reset(int col_num, int maxRow)
        {
            for (int i = 2; i < maxRow; i++)
            {
                string str = worksheet.Cell(i, col_num).Value.ToString().Trim();
                worksheet.Cell(i, col_num).Value = str;
                string[] strArray = str.Trim().Split(' ');
                str = "";
                for (int str_c = 0; str_c < strArray.Length; str_c++)
                {
                    if (strArray[str_c] != "")
                    {
                        str += strArray[str_c];
                        if (str_c + 1 < strArray.Length)
                        {
                            str += "$";
                        }
                    }
                }
                worksheet.Cell(i, col_num).Value = str;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            int checknum = 0;
            workbook = new XLWorkbook(txtExcelPath.Text);
            worksheet = workbook.Worksheet(1);
            int maxRows = 1000;
            XLWorkbook newWorkBook = new XLWorkbook();
            IXLWorksheet newWorkSheet = newWorkBook.Worksheets.Add("plm");
            newWorkSheet.Cell(1, 1).Value = "主料選型代碼";
            newWorkSheet.Cell(1, 2).Value = "位數";
            newWorkSheet.Cell(1, 3).Value = "規格";
            newWorkSheet.Cell(1, 4).Value = "編碼代號";
            newWorkSheet.Cell(1, 5).Value = "編碼標籤";
            newWorkSheet.Cell(1, 6).Value = "起始位數";
            newWorkSheet.Cell(1, 7).Value = "長度";
            int newSheetRow = 2;


            string changeOther = "";
            for (int i = 2; i < maxRows; i++)
            {
                string col_item_number = worksheet.Cell(i, 4).Value.ToString();
                string[] col_Array = col_item_number.Split('$');
                string col1= worksheet.Cell(i, 1).Value.ToString();
                string col2 = worksheet.Cell(i, 2).Value.ToString();
                string col3 = worksheet.Cell(i, 3).Value.ToString();
                string[] arr_col1 = col1.Split(' ');
                if(changeOther!= arr_col1[0])
                {
                    newWorkSheet.Cell(newSheetRow, 1).Value = arr_col1[0];
                    newWorkSheet.Cell(newSheetRow, 3).Value = "類別";
                    newWorkSheet.Cell(newSheetRow, 4).Value = arr_col1[0];
                    newWorkSheet.Cell(newSheetRow, 5).Value = arr_col1[0];
                    newWorkSheet.Cell(newSheetRow, 6).Value = 1;
                    newWorkSheet.Cell(newSheetRow, 7).Value = 4;
                    newSheetRow += 1;
                    newWorkSheet.Cell(newSheetRow, 1).Value = arr_col1[0];
                    newWorkSheet.Cell(newSheetRow, 3).Value = "區分";
                    newWorkSheet.Cell(newSheetRow, 4).Value = "-";
                    newWorkSheet.Cell(newSheetRow, 5).Value = "-";
                    newWorkSheet.Cell(newSheetRow, 6).Value = 5;
                    newWorkSheet.Cell(newSheetRow, 7).Value = 1;
                    newSheetRow += 1;
                    changeOther = arr_col1[0];
                }
                
                
                for (int j = 0; j < col_Array.Length; j++)
                {
                    
                    if (col2.IndexOf('─') > -1)
                    {
                        string[] col2_num = col2.Split('─');
                        regex_string = col2_num[0];
                        if(int.TryParse(col2_num[0],out checknum) )
                        {
                            int firstNum = int.Parse(RegexString(col2_num[0]));
                            int thirdNum = int.Parse(RegexString(col2_num[1]));
                            int lenth = thirdNum - firstNum + 1;
                            newWorkSheet.Cell(newSheetRow, 6).Value = firstNum;
                            newWorkSheet.Cell(newSheetRow, 7).Value = lenth;
                        }
                    }else
                    {
                        if (int.TryParse(col2, out checknum)==false)
                        {
                            continue;
                        }
                        newWorkSheet.Cell(newSheetRow, 6).Value = col2;
                        newWorkSheet.Cell(newSheetRow, 7).Value = 1;
                    }
                    newWorkSheet.Cell(newSheetRow, 1).Value = arr_col1[0];
                    newWorkSheet.Cell(newSheetRow, 2).Value = col2;
                    newWorkSheet.Cell(newSheetRow, 3).Value = col3;
                    if (col_Array[j].IndexOf("=") > -1)
                    {
                        string[] equal = col_Array[j].Split('=');
                        if (equal.Length > 1)
                        {
                            newWorkSheet.Cell(newSheetRow, 4).Value = equal[0];
                            newWorkSheet.Cell(newSheetRow, 5).Value = equal[1];
                        }else
                        {
                            newWorkSheet.Cell(newSheetRow, 4).Value = equal[0];
                        }
                    }else
                    {
                        newWorkSheet.Cell(newSheetRow, 4).Value = col_Array[j];
                    }
                    newSheetRow += 1;
                }
            }

            newWorkSheet.Columns().Width=10;
            newWorkBook.SaveAs(txtResult.Text);
        }
        private static string RegexString(string str)
        {
            Regex rgx = new Regex(regex_condition);
            return rgx.Replace(str, "");
        }
    }
}
