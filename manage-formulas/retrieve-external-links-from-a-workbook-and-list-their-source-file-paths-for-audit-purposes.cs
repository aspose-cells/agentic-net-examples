// Title: Identify and list external workbook references in an Excel file using Aspose.Cells for .NET (C#)
// Common Searches: how to find external workbook references in an Excel file with Aspose.Cells C# | C# code to enumerate external links and their source file paths in a .xlsx using Aspose.Cells | Aspose.Cells detect formulas that reference other workbooks in .NET | list cells containing external links in a workbook with Aspose.Cells for .NET | audit external formula references in Excel using the Aspose.Cells library in C#
// Tags: enumerate external workbook links Aspose.Cells | detect external formula references .NET | list source file paths from Excel formulas C# | audit external links in .xlsx using Aspose.Cells | iterate worksheets cells Aspose.Cells C#

using System;
using System.IO;
using Aspose.Cells;

// The example verifies the presence of an input .xlsx file, loads it with Aspose.Cells, iterates through each worksheet and cell, checks formulas for bracketed workbook names (e.g., [Book1.xlsx]Sheet1!A1), and prints the worksheet name, cell address, and formula for every external reference found, reporting when none exist and handling runtime exceptions.
class ExternalLinksAudit
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);
            bool anyLinksFound = false;

            // Iterate through each worksheet and inspect formulas for external references
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                bool sheetHasLinks = false;

                // Scan all cells in the worksheet
                foreach (Cell cell in sheet.Cells)
                {
                    if (cell.IsFormula)
                    {
                        string formula = cell.Formula;

                        // Simple detection of external workbook reference (e.g., [Book1.xlsx]Sheet1!A1)
                        if (formula.Contains("[") && formula.Contains("]"))
                        {
                            if (!sheetHasLinks)
                            {
                                Console.WriteLine($"External links in worksheet \"{sheet.Name}\":");
                                sheetHasLinks = true;
                            }

                            Console.WriteLine($"Cell {cell.Name}: {formula}");
                            anyLinksFound = true;
                        }
                    }
                }
            }

            if (!anyLinksFound)
            {
                Console.WriteLine("No external links found in the workbook.");
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors during processing
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
