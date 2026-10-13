// Title: C# using Aspose.Cells to replace #DIV/0! formula errors with zero in an Excel workbook
// AI Prompts: Generate C# code with Aspose.Cells that scans every worksheet, finds cells whose evaluated value is '#DIV/0!' and writes 0 instead. | Show how to clear a formula and assign a constant zero to cells that cause division‑by‑zero errors using Aspose.Cells. | Provide a method that iterates the used range of a workbook and replaces any #DIV/0! error result with a numeric zero in C#.
// Common Searches: asp.net aspose.cells how to change #DIV/0! to 0 in Excel file | c# program to fix division by zero errors in Excel formulas using Aspose | replace Excel #DIV/0! error with zero when processing workbook with Aspose.Cells | detect and correct #DIV/0! cells in all worksheets via C# Aspose.Cells | remove formula that returns #DIV/0! and set cell value to zero in .NET
// Tags: Aspose.Cells replace division by zero error with zero | C# iterate workbook cells formula evaluation | remove formula and write constant zero with Aspose.Cells | handle #DIV/0! error in Excel via Aspose.Cells | Aspose.Cells used range error value substitution

using Aspose.Cells;
using System;
using System.IO;

// The example loads an Excel workbook, iterates through each worksheet's used range, identifies cells whose formulas evaluate to the #DIV/0! error, clears the offending formulas, writes a zero value to those cells, and saves the corrected file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                try
                {
                    Cells cells = sheet.Cells;

                    // Determine the used range to limit iteration
                    int maxRow = cells.MaxDataRow;
                    int maxCol = cells.MaxDataColumn;

                    for (int row = 0; row <= maxRow; row++)
                    {
                        for (int col = 0; col <= maxCol; col++)
                        {
                            Cell cell = cells[row, col];

                            // Process only cells that contain a formula
                            if (!string.IsNullOrEmpty(cell.Formula))
                            {
                                // Accessing cell.Value triggers formula evaluation
                                object value = cell.Value;

                                // Check if the evaluated value is the #DIV/0! error (represented as a string)
                                if (value is string strValue && strValue.Equals("#DIV/0!", StringComparison.OrdinalIgnoreCase))
                                {
                                    // Remove the formula and set the cell value to zero
                                    cell.Formula = string.Empty;
                                    cell.PutValue(0);
                                }
                            }
                        }
                    }
                }
                catch (Exception sheetEx)
                {
                    Console.WriteLine($"Error processing sheet '{sheet.Name}': {sheetEx.Message}");
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
