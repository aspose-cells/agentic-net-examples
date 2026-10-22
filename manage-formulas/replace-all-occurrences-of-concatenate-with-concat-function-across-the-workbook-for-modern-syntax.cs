// Title: Replace CONCATENATE with CONCAT in every formula of an Excel workbook using Aspose.Cells for .NET
// Common Searches: how to programmatically change CONCATENATE to CONCAT in an .xlsx file with Aspose.Cells C# | Aspose.Cells replace legacy CONCATENATE function across all worksheets | C# code to update Excel formulas from CONCATENATE to CONCAT using Aspose.Cells | bulk formula text substitution in Excel workbook with Aspose.Cells .NET | iterate through cells and modify formula strings in Aspose.Cells
// Tags: replace CONCATENATE with CONCAT Aspose.Cells | bulk formula update .xlsx C# | iterate worksheets cells Aspose.Cells | modify Excel formula text programmatically | load save workbook after formula change Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook with Aspose.Cells, walks through each worksheet and cell, detects formulas containing the legacy CONCATENATE function, replaces it with the modern CONCAT function, and saves the modified workbook to a new file.
class Program
{
    static void Main()
    {
        // Paths for input and output workbooks
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook (use LoadOptions in case the file is password‑protected)
            var loadOptions = new LoadOptions
            {
                // If a password is required, set it here; empty string for no password
                Password = string.Empty
            };
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;

                // Iterate through all cells that contain data
                foreach (Cell cell in cells)
                {
                    // Check if the cell contains a formula with CONCATENATE
                    if (cell.IsFormula && cell.Formula != null && cell.Formula.Contains("CONCATENATE"))
                    {
                        // Replace CONCATENATE with CONCAT
                        string updatedFormula = cell.Formula.Replace("CONCATENATE", "CONCAT");
                        cell.Formula = updatedFormula;
                    }
                }
            }

            // Save the updated workbook
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
