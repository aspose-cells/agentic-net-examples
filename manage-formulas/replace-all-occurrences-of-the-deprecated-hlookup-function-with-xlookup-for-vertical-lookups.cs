// Title: Replace HLOOKUP formulas with XLOOKUP across all worksheets using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that loads a workbook, scans every worksheet, finds cells whose formula contains the HLOOKUP function (case‑insensitive), replaces it with XLOOKUP while keeping the original arguments, and saves the workbook to a new file. | Create a version of the script that logs the address, original formula, and updated formula for each replaced cell to the console before saving. | Modify the program to accept input and output file paths as command‑line arguments and perform the HLOOKUP‑to‑XLOOKUP conversion without hard‑coded filenames.
// Common Searches: Aspose.Cells C# replace HLOOKUP with XLOOKUP in all sheets | how to update Excel formulas programmatically using Aspose.Cells .NET | convert deprecated Excel lookup functions to XLOOKUP with C# | iterate through every cell in a workbook and change formulas Aspose.Cells | case‑insensitive search and replace in Excel formulas using Aspose.Cells
// Tags: convert HLOOKUP to XLOOKUP Aspose.Cells | bulk formula replacement Excel .NET | iterate worksheets cells Aspose.Cells C# | case-insensitive formula search Aspose.Cells | save updated workbook Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The program loads an Excel workbook, iterates through all worksheets and used cells, detects formulas containing the deprecated HLOOKUP function (case‑insensitive), replaces each occurrence with XLOOKUP while preserving the original arguments, and saves the modified workbook to a new file, handling missing files and runtime errors.
class ReplaceHlookupWithXlookup
{
    static void Main()
    {
        try
        {
            const string inputPath = "InputWorkbook.xlsx";
            const string outputPath = "OutputWorkbook.xlsx";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Get the cells collection of the current worksheet
                Cells cells = sheet.Cells;

                // Iterate through all used cells to check for formulas
                foreach (Cell cell in cells)
                {
                    // Process only cells that contain a formula
                    if (cell.IsFormula)
                    {
                        string formula = cell.Formula;

                        // Check if the formula uses the deprecated HLOOKUP function
                        if (formula.IndexOf("HLOOKUP", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            // Replace HLOOKUP with XLOOKUP.
                            // Note: This simple replacement keeps the original arguments.
                            // For more accurate conversion, adjust the argument order as needed.
                            string updatedFormula = formula.Replace("HLOOKUP", "XLOOKUP", StringComparison.OrdinalIgnoreCase);
                            cell.Formula = updatedFormula;
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
