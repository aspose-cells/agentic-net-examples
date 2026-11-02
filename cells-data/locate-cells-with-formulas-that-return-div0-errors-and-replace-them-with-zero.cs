// Title: Replace #DIV/0! formula errors with zero values in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# replace #DIV/0! error values with 0 in Excel workbook | detect division by zero formula errors and set them to zero using Aspose.Cells | iterate used range cells and fix #DIV/0! errors in .NET Excel processing | calculate formulas then replace error results with zero in Aspose.Cells C# | remove #DIV/0! cells from an Excel file programmatically with Aspose.Cells
// Tags: replace division by zero errors Aspose.Cells | formula error value substitution .NET | iterate used range cells Aspose.Cells | calculate workbook formulas before inspection | set cell value to zero for error results C#

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel file with Aspose.Cells, forces formula calculation, scans each worksheet's used range, identifies cells whose formulas evaluate to the "#DIV/0!" error, replaces those error values with zero, and saves the corrected workbook.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Ensure all formulas are calculated so we can inspect their results
            workbook.CalculateFormula();

            // Iterate through each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;

                // Determine the used range to avoid scanning empty cells
                int maxRow = cells.MaxDataRow;
                int maxCol = cells.MaxDataColumn;

                // Scan each cell in the used range
                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        Cell cell = cells[row, col];

                        // Process only cells that contain a formula
                        if (cell.IsFormula)
                        {
                            // After calculation, errors are typically represented as strings like "#DIV/0!"
                            if (cell.Value is string s && s.StartsWith("#"))
                            {
                                // Replace the error result with zero
                                cell.PutValue(0);
                            }
                        }
                    }
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
