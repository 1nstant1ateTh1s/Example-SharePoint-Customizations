using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.IO;
using Microsoft.SharePoint;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;

namespace RFQEventReceiver
{
    static public class ExcelDocumentUtil
    {
        /// <summary>
        /// Extracts & returns the Excel data within an Excel file as a DataTable object.
        /// </summary>
        /// <param name="file">The uploaded Excel file.</param>
        /// <returns>A table object containing the excel file data.</returns>
        static public DataTable ExtractExcelSheetToDataTable(SPFile file)
        {
            // Variables
            DataTable dt = new DataTable();
            Stream dataStream;
            string fileExt = Path.GetExtension(file.Name);
            if (fileExt.ToLower() != ".xlsx")
            {

                // Provided File is of the wrong file type (Must be an Excel '07 + file)
                throw new Exception("Excel file was not recognized.");
            }

            try
            {
                // Create binary stream for opening the file
                dataStream = file.OpenBinaryStream();

                // Open the spreadsheet document w/ read-only access
                using (SpreadsheetDocument doc =
                    SpreadsheetDocument.Open(dataStream, false))
                {
                    // Retrieve references
                    WorkbookPart wbPart = doc.WorkbookPart;
                    WorksheetPart wsPart = wbPart.WorksheetParts.LastOrDefault(); // multiple worksheets are in decending order - grabbing the last position will retrieve the "first" sheet
                    SharedStringTablePart ssTblPart = wbPart.SharedStringTablePart;
                    SheetData sheetData = wsPart.Worksheet.GetFirstChild<SheetData>();

                    // Ensure worksheet part exists
                    if (wsPart != null)
                    {
                        // Obtain reference to the first & last rows of the worksheet         
                        Row firstRow = sheetData.Elements<Row>().FirstOrDefault();

                        // LINQ query to skip first row w/ column names & grab all remaining rows
                        IEnumerable<Row> dataRows =
                            from row in sheetData.Elements<Row>()
                            where row.RowIndex > 1
                            select row;

                        // Check for first row w/ column names
                        if (firstRow != null)
                        {
                            // Loop through first row's cells to grab column names
                            var cells = firstRow.Descendants<Cell>();
                            foreach (Cell c in cells)
                            {
                                string value = GetCellValue(c, ssTblPart); // retrieve cell's value ...
                                dt.Columns.Add(value); // ... & use to add new column to object
                            }
                        }

                        // Check for the remaining rows w/ the data
                        foreach (Row row in dataRows)
                        {
                            // LINQ query to return the row's cell values 
                            var cells = row.Descendants<Cell>();
                            IEnumerable<string> cellValues =
                                from cell in cells
                                select (GetCellValue(cell, ssTblPart));
                                //select (Convert.ToString(getCellValue(cell, wbPart)));

                            // Check to verify that the row contained data
                            if (cellValues.Count() > 0)
                            {
                                // Start a new row w/ a schema based on the table object
                                DataRow dr = dt.NewRow();
                                dr.ItemArray = cellValues.ToArray(); // transfer spreadsheet cell values to the new data row
                                dt.Rows.Add(dr); // add row to our table object
                            }
                            else
                            {
                                // If no cells, then we have reached the end of the table
                                break;
                            }
                        }
                    }
                }

                // return data object containing the Excel data
                return dt;
            }
            catch (Exception ex)
            {
                throw ex;
            }

            //return dt;
        }

        /// <summary>
        /// Returns the value of a cell. If the cell's value is stored in a shared string table, the value is looked up & returned.
        /// </summary>
        /// <param name="cell">The Excel Spreadsheet cell to reference.</param>
        /// <param name="ssTblPart">The SharedStringTablePart of the Workbook, in case the cell's value is a shared string.</param>
        /// <returns>The string value of the cell.</returns>
        static public string GetCellValue(Cell cell, SharedStringTablePart ssTblPart)
        {
            // Variables
            SharedStringTable sharedStringTbl = ssTblPart.SharedStringTable;
            string value = null;

            // Return the value of a cell unless the cell is empty.
            // If the cell contains a Shared String, its value will be a reference id which will be used to look up the value in the 
            // Shared String table.
            if (cell != null && cell.ChildElements.Count > 0)
            {
                value = ((cell.DataType != null &&
                    cell.DataType.Value == CellValues.SharedString)
                    ? (sharedStringTbl.ChildElements[int.Parse(cell.CellValue.InnerText)].InnerText)
                    : (cell.CellValue != null ? cell.CellValue.InnerText : ""));
            }

            // Return cell value
            return value;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="val"></param>
        /// <returns></returns>
        static public DateTime ReadExcelDateTimeValue(object val)
        {
            DateTime dt;
            if (!(DateTime.TryParse(val.ToString(), out dt))) // attempt to parse string to DateTime object ...
            {
                dt = DateTime.FromOADate(Convert.ToDouble(val)); // ... or convert Excel's Julian date format to .NET DateTime object
            }
            return dt; // return DateTime value
        }
    }
}
