// Title: How to remove an Excel table totals row and re‑add it with custom SUM formulas for numeric columns using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# remove table totals row and add SUM formulas | C# hide totals row then show it again in an Excel table with Aspose.Cells | Set structured reference SUM formula for each numeric column in an Aspose.Cells ListObject | Detect numeric columns in an Excel table and apply totals row formulas using Aspose.Cells .NET | Recreate an Excel table totals row after deletion with Aspose.Cells for .NET
// Tags: Aspose.Cells remove table totals row C# | Aspose.Cells add SUM formulas to totals row | structured reference formula in Aspose.Cells ListObject | detect numeric columns in Excel table Aspose.Cells | recreate Excel table totals row Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The example loads an Excel workbook, checks for a ListObject, hides and then shows the totals row to recreate it, calculates the totals row index, iterates each column to determine if the first data cell is numeric, applies a structured‑reference SUM formula to numeric columns, clears non‑numeric totals cells, and saves the updated workbook.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify input file exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        Workbook workbook;
        try
        {
            // Load the workbook
            workbook = new Workbook(inputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to load workbook: {ex.Message}");
            return;
        }

        try
        {
            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure there is at least one table (ListObject) in the worksheet
            if (worksheet.ListObjects.Count == 0)
            {
                Console.WriteLine("No tables (ListObjects) found in the worksheet.");
                return;
            }

            ListObject table = worksheet.ListObjects[0];

            // Hide and then show totals row to (re)create it
            table.ShowTotals = false;
            table.ShowTotals = true;

            // Calculate the index of the totals row
            // Header row + data rows + totals row
            int totalRowIndex = table.StartRow + 1 + table.DataRange.RowCount;

            // Iterate through each column in the table
            for (int colIndex = 0; colIndex < table.ListColumns.Count; colIndex++)
            {
                int dataRow = table.StartRow + 1; // first data row (after header)
                int dataCol = table.StartColumn + colIndex;
                Cell firstDataCell = worksheet.Cells[dataRow, dataCol];

                // Determine if the first data cell contains a numeric value
                bool isNumeric = firstDataCell.Value is double ||
                                 firstDataCell.Value is int ||
                                 firstDataCell.Value is decimal;

                // Reference the cell in the totals row for this column
                Cell totalCell = worksheet.Cells[totalRowIndex, dataCol];

                if (isNumeric)
                {
                    // Build SUM formula for the column using structured reference
                    string columnName = table.ListColumns[colIndex].Name;
                    string formula = $"=SUM([{columnName}])";

                    // Apply the formula to the totals row cell
                    totalCell.Formula = formula;
                }
                else
                {
                    // Clear any existing formula or value for non‑numeric columns
                    totalCell.Formula = string.Empty;
                    totalCell.PutValue(string.Empty);
                }
            }

            // Save the modified workbook
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred during processing: {ex.Message}");
        }
    }
}
