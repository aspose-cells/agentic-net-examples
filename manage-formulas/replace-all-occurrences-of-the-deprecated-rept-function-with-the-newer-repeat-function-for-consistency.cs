// Title: C# code to replace the deprecated REPT function with REPEAT in Excel formulas using Aspose.Cells
// Common Searches: C# Aspose.Cells replace REPT with REPEAT in workbook formulas | How to update Excel formulas from REPT to REPEAT programmatically with Aspose.Cells | Iterate through cells and change deprecated REPT function to REPEAT using Aspose.Cells .NET | Replace deprecated Excel function REPT in all worksheets with Aspose.Cells API | Aspose.Cells formula editing replace old function with new function in .NET
// Tags: Aspose.Cells formula manipulation C# | Excel formula function migration .NET | update deprecated Excel functions using Aspose.Cells | programmatic Excel cell formula editing | Aspose.Cells workbook formula replacement

using Aspose.Cells;
using System;
using System.IO;

// The example loads an Excel workbook with Aspose.Cells, walks through each worksheet and cell, detects formulas that contain the obsolete REPT function, replaces each occurrence with the newer REPEAT function, and saves the updated workbook to a new file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists before attempting to load it
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook (no password is supplied; adjust LoadOptions if a password is required)
            var loadOptions = new LoadOptions(); // add password here if needed: loadOptions.Password = "yourPassword";
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;

                // Loop through all cells that contain formulas
                foreach (Cell cell in cells)
                {
                    if (cell.IsFormula)
                    {
                        string formula = cell.Formula;
                        // Replace deprecated REPT function with REPEAT, if present
                        if (!string.IsNullOrEmpty(formula) && formula.Contains("REPT("))
                        {
                            cell.Formula = formula.Replace("REPT(", "REPEAT(");
                        }
                    }
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (CellsException ex)
        {
            Console.WriteLine($"Aspose.Cells error: {ex.Message}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
