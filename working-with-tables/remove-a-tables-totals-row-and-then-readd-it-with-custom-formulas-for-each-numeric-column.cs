// Title: Hide an Excel table’s totals row and re‑add it with structured SUM formulas for numeric columns using Aspose.Cells for .NET
// AI Prompts: Generate C# code with Aspose.Cells that hides a ListObject’s totals row, then shows it again and inserts a SUM formula only for columns that contain numbers. | Create a method that iterates through each column of a worksheet table, detects numeric data types, and assigns a structured reference SUM formula to the totals row. | Write a script that loads an Excel file, validates the presence of a table, toggles the ShowTotals property, and applies custom formulas to the totals row based on column content.
// Common Searches: aspnet hide table totals row and add sum formulas with Aspose.Cells | how to use structured references for totals row in Aspose.Cells C# | detect numeric columns in Excel table and set SUM formula programmatically | Aspose.Cells toggle ShowTotals and apply custom formulas to totals row
// Tags: Aspose.Cells hide table totals row | Aspose.Cells add structured SUM to totals row | detect numeric columns ListObject C# | ShowTotals property Aspose.Cells | Excel table totals row custom formulas

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The program loads an Excel workbook, ensures a table exists, hides its totals row, re‑shows it, scans each column to determine if it holds numeric data, and inserts a structured reference SUM formula into the totals row for those columns before saving the file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure at least one table exists
            if (worksheet.ListObjects.Count == 0)
            {
                Console.WriteLine("No tables found on the worksheet.");
                return;
            }

            // Work with the first table
            ListObject table = worksheet.ListObjects[0];

            // 1. Remove existing totals row
            table.ShowTotals = false;

            // 2. Re‑add totals row
            table.ShowTotals = true;

            // Totals row index after ShowTotals = true
            int totalsRowIndex = table.EndRow;

            // Iterate through each column in the table
            for (int i = 0; i < table.ListColumns.Count; i++)
            {
                ListColumn column = table.ListColumns[i];

                // Worksheet column index (zero‑based)
                int worksheetColumnIndex = table.StartColumn + i;

                // Determine data range rows (excluding header and totals)
                int firstDataRow = table.DataRange.FirstRow;
                int lastDataRow = firstDataRow + table.DataRange.RowCount - 1;

                bool isNumeric = false;

                // Scan cells to detect numeric data
                for (int row = firstDataRow; row <= lastDataRow; row++)
                {
                    object cellValue = worksheet.Cells[row, worksheetColumnIndex].Value;
                    if (cellValue is double || cellValue is int || cellValue is decimal)
                    {
                        isNumeric = true;
                        break;
                    }
                }

                if (isNumeric)
                {
                    // Set SUM formula using structured reference
                    string formula = $"=SUM([{column.Name}])";
                    worksheet.Cells[totalsRowIndex, worksheetColumnIndex].Formula = formula;
                }
                else
                {
                    // Clear any existing formula for non‑numeric columns
                    worksheet.Cells[totalsRowIndex, worksheetColumnIndex].Formula = string.Empty;
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
