// Title: Insert a new row into an Excel ListObject and automatically recalculate column formulas with Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells insert new row into ListObject and keep existing formulas | how to expand an Excel table after adding a physical row with Aspose.Cells | recalculate column formulas automatically after adding data to a ListObject in .NET | retrieve calculated total value from a newly added row using Aspose.Cells
// Tags: insert row and resize ListObject Aspose.Cells | auto recalculate formulas after table row addition Aspose.Cells | get column index by header name C# Aspose.Cells | populate ListObject columns programmatically Aspose.Cells | save updated workbook to new file Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Tables;
using System;
using System.IO;

// The example loads a workbook, locates the ListObject named "SalesTable", inserts a physical row at the end of the table, expands the ListObject to include the new row, fills the Product, Quantity, and UnitPrice columns, triggers a full formula recalculation, optionally reads the calculated Total value, and saves the modified workbook to a new file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input file exists.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook.
            Workbook workbook = new Workbook(inputPath);
            Worksheet sheet = workbook.Worksheets[0];

            // Get the table named "SalesTable".
            ListObject salesTable = sheet.ListObjects["SalesTable"];
            if (salesTable == null)
            {
                Console.WriteLine("Table 'SalesTable' not found.");
                return;
            }

            // Determine the index of the new row (after the current data rows).
            int newRowIndex = salesTable.DataRange.FirstRow + salesTable.DataRange.RowCount;

            // Insert a physical row in the worksheet.
            sheet.Cells.InsertRows(newRowIndex, 1);

            // Expand the table to include the new row (hasHeaders = true).
            salesTable.Resize(
                salesTable.DataRange.FirstRow,
                salesTable.DataRange.FirstColumn,
                salesTable.DataRange.RowCount + 1,
                salesTable.DataRange.ColumnCount,
                true);

            // Helper to find column index by header name.
            int GetColumnIndex(string columnName)
            {
                int headerRow = salesTable.DataRange.FirstRow - 1; // header is just above data range
                int firstCol = salesTable.DataRange.FirstColumn;
                int lastCol = firstCol + salesTable.DataRange.ColumnCount - 1;

                for (int col = firstCol; col <= lastCol; col++)
                {
                    if (sheet.Cells[headerRow, col].StringValue.Equals(columnName, StringComparison.OrdinalIgnoreCase))
                        return col;
                }

                throw new ArgumentException($"Column '{columnName}' not found in table.");
            }

            // Set values for the new row.
            int productCol = GetColumnIndex("Product");
            int quantityCol = GetColumnIndex("Quantity");
            int unitPriceCol = GetColumnIndex("UnitPrice");

            sheet.Cells[newRowIndex, productCol].PutValue("New Product");
            sheet.Cells[newRowIndex, quantityCol].PutValue(10);
            sheet.Cells[newRowIndex, unitPriceCol].PutValue(5.5);

            // Recalculate formulas.
            workbook.CalculateFormula();

            // Retrieve the calculated total (if a "Total" column exists).
            try
            {
                int totalCol = GetColumnIndex("Total");
                object totalValue = sheet.Cells[newRowIndex, totalCol].Value;
                Console.WriteLine($"Calculated Total for the new row: {totalValue}");
            }
            catch (ArgumentException)
            {
                // "Total" column not present; ignore.
            }

            // Save the updated workbook.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
