/****************************** Module Header ******************************\
* Module Name:  GetWordPlainText.cs
* Project:      CSOpenXmlGetPlainText
* Copyright(c) Microsoft Corporation.
* 
* The Class is used to read plain text from word document.
* Microsoft Word *.docx is an Open XML document combining texts, stytle,grapyhics 
* and so on into a single ZIP archive. 
* The Class uses Open XML SDK API to read XML element and filter the text. 
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
using System.IO;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ClosedXML.Excel;
using System.Collections;
using System.Collections.Generic;
using System.Data;

namespace CSUsingOpenXmlPlainText
{
    public class GetWordPlainText : IDisposable
    {
        // Specify whether the instance is disposed.
        private bool disposed = false;

        // The word package
        private WordprocessingDocument package = null;
        private XLWorkbook workbook = null;
        private IXLWorksheet worksheet = null;
        private int CoRow = 2;
        /// <summary>
        ///  Get the file name
        /// </summary>
        private string FileName = string.Empty;

        /// <summary>
        ///  Initialize the WordPlainTextManager instance
        /// </summary>
        /// <param name="filepath"></param>
        public GetWordPlainText(string filepath)
        {
            this.FileName = filepath;
            if (string.IsNullOrEmpty(filepath) || !File.Exists(filepath))
            {
                throw new Exception("The file is invalid. Please select an existing file again");
            }

            this.package = WordprocessingDocument.Open(filepath, true);
            workbook = new XLWorkbook();
            worksheet = workbook.Worksheets.Add("plm");
            worksheet.Cell(1, 1).Value = "主料選型代碼";
            worksheet.Cell(1, 2).Value = "位數";
            worksheet.Cell(1, 3).Value = "規格";
            worksheet.Cell(1, 4).Value = "編碼說明";
        }

        /// <summary>
        ///  Read Word Document
        /// </summary>
        /// <returns>Plain Text in document </returns>
        public string ReadWordDocument(string path)
        {
            StringBuilder sb = new StringBuilder();
            OpenXmlElement element = package.MainDocumentPart.Document.Body;
            if (element == null)
            {
                return string.Empty;
            }
            
            sb.Append(GetPlainText(element));
            worksheet.Columns().AdjustToContents();
            workbook.SaveAs(path);

            return sb.ToString();
        }

        /// <summary>
        ///  Read Plain Text in all XmlElements of word document
        /// </summary>
        /// <param name="element">XmlElement in document</param>
        /// <returns>Plain Text in XmlElement</returns>
        public string GetPlainText(OpenXmlElement element)
        {
            StringBuilder PlainTextInWord = new StringBuilder();
            foreach (OpenXmlElement section in element.Elements())
            {
                switch (section.LocalName)
                {
                    case "p":
                        //PlainTextInWord.Append(GetPlainText(section));
                        foreach (OpenXmlElement section_p in section.Elements())
                        {
                            switch (section_p.LocalName)
                            {
                                case "r":
                                    foreach (OpenXmlElement section_r in section_p.Elements())
                                    {
                                        switch (section_r.LocalName)
                                        {
                                            case "t":
                                                if (section_r.InnerText.IndexOf("52") > -1 && section_r.InnerText.IndexOf("規格")<0)
                                                {
                                                    worksheet.Cell(CoRow, 1).Value = section_r.InnerText;
                                                    //CoRow += 1;
                                                }
                                                break;
                                        }
                                    }
                                    break;
                            }
                        }
                        break;
                    case "tbl":
                        string str = GetTableData(section);
                        break;
                        
                }
            }

            return PlainTextInWord.ToString();
        }
        private string GetTableData(OpenXmlElement element_table)
        {
            bool title = true;
            StringBuilder PlainTextInWord = new StringBuilder();
            string number = "";
            foreach (OpenXmlElement section in element_table.Elements())
            {
                    switch (section.LocalName)
                    {
                        case "tr":
                            int Col = 2;
                            if (title == true)
                            {
                                number = worksheet.Cell(CoRow, 1).Value.ToString();
                                title = false;
                            }else
                            {
                                foreach (OpenXmlElement section_tc in section.Elements())
                                {
                                    switch (section_tc.LocalName)
                                    {
                                        case "tc":
                                            worksheet.Cell(CoRow, 1).Value = number;
                                            worksheet.Cell(CoRow, Col).Value = section_tc.InnerText;
                                            worksheet.Cell(CoRow, Col).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                                            worksheet.Cell(CoRow, Col).Style.Border.OutsideBorderColor = XLColor.Black;
                                            Col += 1;
                                            PlainTextInWord.Append(section_tc.InnerText);
                                            break;
                                    }
                                }
                                CoRow += 1;
                            }
                            
                            break;
                    }
                
            }
            return PlainTextInWord.ToString();
        }
        #region IDisposable interface

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            // Protect from being called multiple times.
            if (disposed)
            {
                return;
            }

            if (disposing)
            {
                // Clean up all managed resources.
                if (this.package != null)
                {
                    this.package.Dispose();
                }
            }

            disposed = true;
        }
        #endregion
    }
}
