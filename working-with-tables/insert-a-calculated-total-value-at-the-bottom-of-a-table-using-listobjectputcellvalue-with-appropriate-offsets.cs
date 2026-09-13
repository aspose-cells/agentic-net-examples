// Title: Add a calculated total row below an Aspose.Cells ListObject table using PutCellValue in C#
// AI Prompts: Generate C# code that appends a "Total" label and a SUM formula for a specific column after an Aspose.Cells ListObject, using PutCellValue with row and column offsets. | Show how to compute and insert a column total in an Excel table created with Aspose.Cells, then trigger workbook.CalculateFormula to display the result.
// Common Searches: how to append a total row after a ListObject in Aspose.Cells C# | using PutCellValue to write a SUM formula below an Excel table with Aspose.Cells | calculate column sum for a ListObject table programmatically in C# | force formula evaluation after inserting totals in Aspose.Cells workbook | determine row offset for total row in Aspose.Cells ListObject
// Tags: Aspose.Cells ListObject PutCellValue | add total row to Excel table | write SUM formula with Aspose.Cells | row offset calculation for ListObject | force formula calculation C# workbook | C# Excel table total automation

using Aspose.Cells;
using Aspose.Cells.Tables;
using System;

// The example creates a workbook, defines a ListObject named SalesTable with sample data, calculates the correct row and column offsets, uses ListObject.PutCellValue to write a "Total" label and a SUM formula for the Quantity column, forces formula evaluation, and saves the result as TableWithTotal.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Data";

            // Sample data for the table (header + 3 rows)
            object[,] data = new object[,]
            {
                { "Item", "Quantity", "Price" },
                { "Apple", "10", "0.5" },
                { "Banana", "5", "0.3" },
                { "Orange", "8", "0.4" }
            };

            // Populate the worksheet starting at cell A1 (manual import to avoid missing API)
            int rows = data.GetLength(0);
            int cols = data.GetLength(1);
            for (int r = 0; r < rows; r++)
            {
                for (int c = 0; c < cols; c++)
                {
                    sheet.Cells[r, c].PutValue(data[r, c]);
                }
            }

            // Define the range that will become the ListObject (table)
            int firstRow = 0;
            int firstCol = 0;
            int lastRow = rows - 1;   // includes header row
            int lastCol = cols - 1;

            // Add a ListObject covering the data range (including header)
            int listObjectIndex = sheet.ListObjects.Add(firstRow, firstCol, lastRow, lastCol, true);
            ListObject table = sheet.ListObjects[listObjectIndex];
            table.DisplayName = "SalesTable";

            // ------------------------------------------------------------
            // Insert a calculated total row below the table using PutCellValue
            // ------------------------------------------------------------

            // Row offset: one row after the last data row of the table
            int totalRow = table.DataRange.FirstRow + table.DataRange.RowCount + 1; // +1 for the total row
            int totalColLabel = table.DataRange.FirstColumn;               // column A for the "Total" label
            int totalColValue = table.DataRange.FirstColumn + 1;           // column B for the summed quantity

            // Write the "Total" label
            table.PutCellValue(totalRow, totalColLabel, "Total");

            // Build the SUM formula for the Quantity column (column B)
            int quantityColIndex = table.DataRange.FirstColumn + 1;

            // Address of first quantity cell (row after header)
            string startAddress = CellsHelper.CellIndexToName(table.DataRange.FirstRow + 1, quantityColIndex);
            // Address of last quantity cell (last data row)
            string endAddress = CellsHelper.CellIndexToName(table.DataRange.FirstRow + table.DataRange.RowCount, quantityColIndex);
            string sumFormula = $"=SUM({startAddress}:{endAddress})";

            // Write the formula into the total value cell
            table.PutCellValue(totalRow, totalColValue, sumFormula);

            // Force calculation so the total appears immediately
            workbook.CalculateFormula();

            // Save the workbook
            workbook.Save("TableWithTotal.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
