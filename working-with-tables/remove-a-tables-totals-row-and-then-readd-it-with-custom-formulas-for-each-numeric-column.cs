// Title: How to remove and recreate an Excel table totals row with custom SUM formulas using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# hide table totals row and add custom SUM formulas | replace Excel table totals row with structured reference formulas using Aspose.Cells | C# code to set totals row label and numeric column formulas in an Excel ListObject | remove existing totals row and re‑enable it with custom calculations in Aspose.Cells | apply number format to totals row cells in Aspose.Cells C# example
// Tags: table totals row manipulation Aspose.Cells | custom totals row formulas ListObject | structured reference SUM formula C# | numeric column totals number format | re‑enable totals row Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The example loads an Excel workbook, checks for a ListObject on the first worksheet, hides the existing totals row, re‑enables it, and then adds custom SUM formulas (using structured references) to each numeric column while applying a two‑decimal number format. Non‑numeric columns receive a "Total" label or are left blank. The modified workbook is saved to a new file.
class Program
{
    static void Main()
    {
        try
        {
            // Input and output file paths
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one table (ListObject)
            if (sheet.ListObjects.Count == 0)
            {
                Console.WriteLine("No tables (ListObjects) found on the first worksheet.");
                return;
            }

            // Assume the first table is the target
            ListObject table = sheet.ListObjects[0];

            // -------------------------------------------------
            // 1. Remove the existing totals row (if it exists)
            // -------------------------------------------------
            if (table.ShowTotals)
            {
                // Hiding the totals row removes it from the table
                table.ShowTotals = false;
            }

            // -------------------------------------------------
            // 2. Re‑add the totals row with custom formulas
            // -------------------------------------------------
            // Enable the totals row again
            table.ShowTotals = true;

            // Determine the row index of the totals row (first row after the data range)
            int totalsRowIndex = table.DataRange.FirstRow + table.DataRange.RowCount;

            // Iterate through each column in the table
            for (int i = 0; i < table.ListColumns.Count; i++)
            {
                ListColumn column = table.ListColumns[i];

                // Column index in the worksheet (absolute)
                int columnIndex = table.DataRange.FirstColumn + i;

                // Check the type of the first data cell in this column to decide if it's numeric
                Cell firstDataCell = sheet.Cells[table.DataRange.FirstRow, columnIndex];

                if (firstDataCell.Type == CellValueType.IsNumeric)
                {
                    // Custom formula: SUM of the column using structured reference
                    string formula = $"=SUM([{column.Name}])";

                    // Set the formula in the totals row cell
                    Cell totalCell = sheet.Cells[totalsRowIndex, columnIndex];
                    totalCell.Formula = formula;

                    // Apply number format (e.g., two decimal places)
                    Style style = totalCell.GetStyle();
                    style.Number = 10; // "#,##0.00"
                    totalCell.SetStyle(style);
                }
                else
                {
                    // For non‑numeric columns, place a label in the first column
                    Cell totalCell = sheet.Cells[totalsRowIndex, columnIndex];
                    if (i == 0)
                    {
                        totalCell.PutValue("Total");
                    }
                    else
                    {
                        totalCell.PutValue(string.Empty);
                    }
                }
            }

            // -------------------------------------------------
            // 3. Save the modified workbook
            // -------------------------------------------------
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
