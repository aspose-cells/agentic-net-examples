// Title: Insert a SUM formula into a specific cell of an Aspose.Cells table using C#
// AI Prompts: Create a new workbook, add a ListObject covering the range A1:C5, and set the formula "=SUM(A2:A5)" in the second data row, second column of the table. | Modify the sample to place an AVERAGE formula in the first data row, third column of the same table using the Cell.Formula property. | Write C# code that creates a worksheet, defines a table, and assigns any custom formula to a cell inside the table's DataRange.
// Common Searches: asp.net aspose.cells set formula in listobject cell c# example | how to add a sum formula to a table cell using Aspose.Cells .NET | c# insert custom formula into Aspose.Cells table data range
// Tags: set cell formula in Aspose.Cells ListObject | insert SUM formula into Aspose.Cells table cell | assign custom formula to Aspose.Cells table data range | Aspose.Cells C# create table and set formula | Cell.Formula usage with Aspose.Cells tables

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsExample
{
    // Demonstrates creating a workbook, adding a ListObject over A1:C5, inserting a SUM formula into cell B2 of the table's data range, and saving the result as TableWithFormula.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Define a range that will become a table (A1:C5)
                // CreateRange(row, column, totalRows, totalColumns) – zero‑based indices
                AsposeRange tableRange = sheet.Cells.CreateRange(0, 0, 5, 3);

                // Build the address string (e.g., "A1:C5") required by ListObjects.Add
                int startRow = tableRange.FirstRow;
                int startCol = tableRange.FirstColumn;
                int endRow = startRow + tableRange.RowCount - 1;
                int endCol = startCol + tableRange.ColumnCount - 1;
                string rangeAddress = $"{CellsHelper.CellIndexToName(startRow, startCol)}:{CellsHelper.CellIndexToName(endRow, endCol)}";

                // Add a table (ListObject) to the worksheet; give it a name and specify that it has headers
                int tableIndex = sheet.ListObjects.Add("Table1", rangeAddress, true);
                ListObject table = sheet.ListObjects[tableIndex];

                // Insert a formula into the second row, second column of the table's data body (cell B2)
                // DataRange represents the body of the table without the header row
                Cell targetCell = table.DataRange[1, 1]; // zero‑based indices within the data range
                targetCell.Formula = "=SUM(A2:A5)";

                // Save the workbook to a file
                string outputPath = "TableWithFormula.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
