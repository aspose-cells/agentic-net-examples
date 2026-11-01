// Title: Replace HLOOKUP formulas with XLOOKUP in an Excel workbook using Aspose.Cells for .NET
// Common Searches: how to replace HLOOKUP with XLOOKUP in a workbook using Aspose.Cells C# | Aspose.Cells bulk update of Excel formulas from HLOOKUP to XLOOKUP | iterate through all worksheets and cells to change HLOOKUP to XLOOKUP in .NET | programmatically convert deprecated HLOOKUP functions to XLOOKUP with Aspose.Cells | C# replace HLOOKUP formulas across an entire Excel file using Aspose API
// Tags: replace HLOOKUP with XLOOKUP Aspose.Cells | bulk formula modification .NET | iterate worksheets cells Aspose.Cells | update deprecated Excel functions C# | programmatic Excel formula conversion Aspose

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook with Aspose.Cells, checks for password protection, iterates through every worksheet and cell, replaces any HLOOKUP function in cell formulas with XLOOKUP (case‑insensitive), and saves the updated workbook to a new file.
class ReplaceHlookupWithXlookup
{
    static void Main()
    {
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            Workbook workbook;

            try
            {
                // Load the workbook (no password). If the file is password‑protected,
                // Aspose.Cells will throw a CellsException.
                workbook = new Workbook(inputPath);
            }
            catch (CellsException ex)
            {
                // Simple detection of a password‑protected file based on the exception message
                if (ex.Message != null && ex.Message.IndexOf("password", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    Console.WriteLine("The workbook is password protected. Please provide a password.");
                }
                else
                {
                    Console.WriteLine($"Failed to load workbook: {ex.Message}");
                }
                return;
            }

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;

                // Iterate through all cells in the worksheet
                foreach (Cell cell in cells)
                {
                    // Replace HLOOKUP with XLOOKUP in formulas (case‑insensitive)
                    if (!string.IsNullOrEmpty(cell.Formula) &&
                        cell.Formula.IndexOf("HLOOKUP", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        string newFormula = cell.Formula.Replace("HLOOKUP", "XLOOKUP", StringComparison.OrdinalIgnoreCase);
                        cell.Formula = newFormula;
                    }
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
