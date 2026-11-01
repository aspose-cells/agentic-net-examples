// Title: How to programmatically add a totals row with SUM formulas to an Excel ListObject using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that loads an .xlsx file, enables the totals row on the first ListObject, and inserts table‑based SUM formulas for every numeric column using Aspose.Cells. | Show me a complete Aspose.Cells example that detects numeric columns in a worksheet table and assigns a SUM formula to the corresponding totals‑row cells.
// Common Searches: C# Aspose.Cells add totals row to Excel table and calculate column sums | How to set SUM formulas in the totals row of a ListObject with Aspose.Cells .NET | Detect numeric columns and auto‑populate totals row in an Excel workbook using Aspose.Cells
// Tags: Aspose.Cells enable ListObject totals row | Aspose.Cells SUM formula for table column | auto detect numeric columns Aspose.Cells | C# write totals row formulas in Excel table | use table display name in Aspose.Cells formula

using Aspose.Cells;
using Aspose.Cells.Tables;
using System;
using System.IO;

// The sample loads an existing workbook, checks for a ListObject on the first worksheet, turns on the table's totals row, computes the totals‑row index, iterates through each column, identifies numeric columns by inspecting the first data cell, builds a SUM formula that references the table's display name and column name, writes the formula into the appropriate totals‑row cell, and saves the updated file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one table (ListObject)
            if (sheet.ListObjects.Count > 0)
            {
                // Work with the first table in the worksheet
                ListObject table = sheet.ListObjects[0];

                // Show the totals row for the table
                table.ShowTotals = true;

                // Calculate the index of the totals row (row after the data range)
                int totalsRowIndex = table.DataRange.FirstRow + table.DataRange.RowCount;

                // Iterate through each column in the table
                for (int col = 0; col < table.ListColumns.Count; col++)
                {
                    ListColumn column = table.ListColumns[col];

                    // Check the type of the first data cell in the column to decide if it's numeric
                    Cell firstDataCell = sheet.Cells[table.DataRange.FirstRow, table.DataRange.FirstColumn + col];
                    if (firstDataCell.Type == CellValueType.IsNumeric)
                    {
                        // Build a SUM formula that references the column by name within the table
                        // Example: =SUM(Table1[Amount])
                        string tableName = table.DisplayName; // Use DisplayName instead of Name
                        string formula = $"=SUM({tableName}[{column.Name}])";

                        // Assign the formula to the corresponding cell in the totals row
                        sheet.Cells[totalsRowIndex, table.DataRange.FirstColumn + col].Formula = formula;
                    }
                }
            }
            else
            {
                Console.WriteLine("No tables (ListObjects) found in the first worksheet.");
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
