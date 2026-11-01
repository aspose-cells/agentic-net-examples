// Title: How to replace TODAY() formulas with a fixed date string in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: C# code to freeze TODAY() function in an Excel file with Aspose.Cells | Aspose.Cells replace dynamic TODAY() formula with static date in .NET | Iterate all cells in a workbook and change TODAY() to a constant date using Aspose.Cells | How to convert TODAY() formulas to a specific date string when saving Excel with Aspose.Cells | Remove volatile TODAY() formulas from Excel worksheets programmatically in C#
// Tags: replace TODAY formula Aspose.Cells C# | static date substitution Excel Aspose.Cells | freeze volatile date calculation .NET | iterate worksheets cells Aspose.Cells | save workbook after formula modification Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example loads an Excel workbook with Aspose.Cells, scans every worksheet and cell for formulas containing TODAY(), replaces each matching formula with a static date string formatted as yyyy-MM-dd, and saves the updated workbook while handling missing files, password protection, and other exceptions.
class Program
{
    static void Main()
    {
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        // Verify that the input file exists before attempting to load it
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        Workbook workbook = null;

        try
        {
            // Attempt to load the workbook (no password supplied)
            workbook = new Workbook(inputPath);
        }
        catch (CellsException ex)
        {
            // Handle password‑protected files or other loading issues
            Console.WriteLine($"Failed to load workbook: {ex.Message}");
            return;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error while loading workbook: {ex.Message}");
            return;
        }

        // Static date string to replace TODAY() formulas
        string staticDate = DateTime.Today.ToString("yyyy-MM-dd");

        try
        {
            // Iterate through each worksheet and replace TODAY() formulas
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;
                int maxRow = cells.MaxDataRow;
                int maxColumn = cells.MaxDataColumn;

                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxColumn; col++)
                    {
                        Cell cell = cells[row, col];
                        if (cell.IsFormula)
                        {
                            string formula = cell.Formula;
                            if (!string.IsNullOrEmpty(formula) &&
                                formula.IndexOf("TODAY()", StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                // Replace the formula with the static date value
                                cell.PutValue(staticDate);
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
            Console.WriteLine($"Error processing workbook: {ex.Message}");
        }
    }
}
