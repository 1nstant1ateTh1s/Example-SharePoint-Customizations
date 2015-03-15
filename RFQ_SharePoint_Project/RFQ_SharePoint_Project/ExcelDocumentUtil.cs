using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace RFQ_SharePoint_Project
{
    static public class ExcelDocumentUtil
    {
        /// <summary>
        /// Given a column name, a row index, & a WorksheetPart, inserts a cell into the worksheet.
        /// If the cell already exists, returns it.
        /// </summary>
        /// <param name="columnName">Reference/name of the target column.</param>
        /// <param name="rowIndex">Index of the target row.</param>
        /// <param name="wsPart">The WorksheetPart to retrieve the Worksheet from.</param>
        /// <returns>A reference to the cell at the specified location.</returns>
        static public Cell InsertCellInWorksheet(string columnName, uint rowIndex, WorksheetPart wsPart)
        {
            // Variables
            Worksheet ws = wsPart.Worksheet;
            SheetData sheetData = ws.GetFirstChild<SheetData>();
            string cellRef = columnName + rowIndex;

            // If the worksheet does not contain a row w/ the specified row index, insert one
            Row row;
            if (sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).Count() != 0)
            {
                row = sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).First();
            }
            else
            {
                row = new Row() { RowIndex = rowIndex };
                sheetData.Append(row);
            }

            // If there is a cell w/ the specified column name + row index, return it
            if (row.Elements<Cell>().Where(c => c.CellReference.Value == columnName + rowIndex).Count() > 0)
            {
                return row.Elements<Cell>().Where(c => c.CellReference.Value == cellRef).First();
            }
            else // if there is not a cell w/ the specified column name + row index, insert one & return it
            {
                Cell refCell = null;
                foreach (Cell cell in row.Elements<Cell>())
                {
                    if (string.Compare(cell.CellReference.Value, cellRef, true) > 0)
                    {
                        refCell = cell;
                        break;
                    }
                }

                // Create the new Cell
                Cell newCell = new Cell() { CellReference = cellRef };
                row.InsertBefore(newCell, refCell);
                
                // Save changes to worksheet & return cell
                ws.Save();
                return newCell;
            }
        }


        /// <summary>
        /// Returns the Excel Column name/identifier (i.e., 'A', 'B', 'Z', 'AA', etc.) for the given property/header/title.
        /// </summary>
        /// <param name="propertyName">The name of the header to look up in row 1 of the Excel worksheet.</param>
        /// <param name="wsPart">The WorkSheetPart of the worksheet to search.</param>
        /// <param name="ssTblPart">The SharedStringTablePart of the Workbook, which contains all of the shared string values.</param>
        /// <returns>The column reference/name for the given property/header/title.</returns>
        static public string GetColumnReference(string propertyName, WorksheetPart wsPart, SharedStringTablePart ssTblPart)
        {
            // Variables
            string colRef = null;
            Worksheet ws = wsPart.Worksheet;
            Row firstRow = wsPart.Worksheet.Descendants<Row>().FirstOrDefault();

            // Attempt to locate the cell containing a value equal to the provided property name
            Cell headerCell = firstRow.Descendants<Cell>().Where(c => GetCellValue(c, ssTblPart) == propertyName).First();
            if (headerCell != null) // cell was found
            {
                // Get the column reference/name for the specified cell.
                // Use a regular expression to match the column reference/name portion of the CellReference.
                string cellRef = headerCell.CellReference;
                string pattern = @"[A-Za-z]+";
                Match match = Regex.Match(cellRef, pattern);
                colRef = match.Value;
            }

            return colRef;
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
                    : (cell.CellValue.InnerText));
            }

            // Return cell's value
            return value;
        }

        /// <summary>
        /// Populates an Excel Cell with the specified value.
        /// </summary>
        /// <param name="cell">The Cell to apply the value to.</param>
        /// <param name="value">The value to apply to the specified Cell.</param>
        /// <param name="ssTblPart">The SharedString's table collection.</param>
        static public void AddValueToCell(Cell cell, object value, SharedStringTablePart ssTblPart)
        {
            // Variables
            Type type = value.GetType();

            // If a cell's value currently comes from the Shared String table, we need to look it up and update the value in that location
            if (cell.DataType != null && cell.DataType == CellValues.SharedString)
            {
                //int text = Int32.Parse(cell.CellValue.Text);
                int text = InsertSharedStringItem(ssTblPart, value.ToString());

                // Assign the new SharedString's position as the CellValue text
                cell.CellValue.Text = text.ToString();
            }
            else // cell does not contain a value that exists within the Shared String table ...
            {
                // If the cell contains an existing value, we can simply overwrite the existing CellValue property w/ the new value; 
                // otherwise, we will need to create the appropriate xml elements
                if (cell.CellValue != null) // cell already contains an existing cell value
                {
                    cell.CellValue.Text = value.ToString(); // just overwrite existing value with new value
                }
                else // cell currently contains no cell value
                {
                    if (type == typeof(System.String)) // new string types will be inserted as "Inline String" elements
                    {
                        // Create an InlineString object to be appended as a child element of the cell
                        InlineString inlineStr = new InlineString();
                        Text cellValueText = new Text() { Text = value.ToString() };
                        inlineStr.AppendChild(cellValueText);
              
                        // Append the InlineString element
                        cell.DataType = CellValues.InlineString;
                        cell.RemoveAllChildren(); // remove any existing "InlineString" elements to prevent any potential conflicts
                        cell.AppendChild(inlineStr);
                    }
                    else // other values such as "Double's" can simply be assigned to the cell's CellValue property to make use of the cell's existing "Style Index"
                    {
                        // Instantiate a new 'CellValue' for this cell
                        cell.CellValue = new CellValue(value.ToString());
                    }
                }
            }
        }

        /// <summary>
        /// Flushes all cells that contain a formula calculated value for a given Spreadsheet Document in order to force their values to refresh.
        /// </summary>
        /// <param name="doc">The Spreadsheet Document on which to flush all of the cells.</param>
        static public void FlushCachedValues(SpreadsheetDocument doc)
        {
            // Flush the value for any cells that contain a formula & an existing value throughout the entire Spreadsheet Document
            doc.WorkbookPart.WorksheetParts
                .SelectMany(part => part.Worksheet.Elements<SheetData>())
                .SelectMany(data => data.Elements<Row>())
                .SelectMany(row => row.Elements<Cell>())
                .Where(cell => cell.CellFormula != null && cell.CellValue != null)
                .ToList()
                .ForEach(cell => cell.CellValue.Remove());
        }

        /// <summary>
        /// Given a SharedStringTablePart and text, creates a SharedStringItem w/ the specified text
        /// & inserts it into the SharedStringTablePart. If the item already exists, returns its index.
        /// </summary>
        /// <param name="ssTblPart">The container for SharedStringItem's.</param>
        /// <param name="value">The text string to find or insert in the SharedStringTablePart.</param>
        /// <returns>The index of the SharedStringItem containing the text value.</returns>
        static private int InsertSharedStringItem(SharedStringTablePart ssTblPart, string value)
        {
            // Variables
            // If the part does not contain a SharedStringTable, create one
            SharedStringTable ssTbl = (ssTblPart.SharedStringTable != null ? ssTblPart.SharedStringTable : new SharedStringTable());
            int position = 0;
            
            // Iterate through all the items in the SharedStringTable. If the text already exists, return its index.
            foreach (SharedStringItem item in ssTbl.Elements<SharedStringItem>())
            {
                if (item.InnerText == value)
                {
                    return position; // return index of the existing SharedStringItem containing the specified text
                }

                position++; // increment counter
            }

            // The text does not exist in the part. Create the SharedStringItem & return its index.
            ssTbl.AppendChild(new SharedStringItem(new Text(value)));
            ssTbl.Save(); // save changes to SharedStringTable

            return position; // return index of newly inserted SharedStringItem
        }
    }
}
