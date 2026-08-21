// Title: Replace GETPIVOTDATA with Structured References using Aspose.Cells for .NET (C#)
// Description: Loads an Excel workbook, scans every worksheet and cell, detects formulas that start with GETPIVOTDATA, substitutes each with a structured‑reference placeholder (e.g., =Table1[Column1]) via Aspose.Cells, and saves the modified file.
// Keywords: Aspose.Cells | C# | .NET | GETPIVOTDATA replacement | structured reference | Excel formula conversion | batch formula update | legacy Excel migration
// Common Searches: Aspose.Cells replace GETPIVOTDATA formula | convert GETPIVOTDATA to table reference C# | update deprecated Excel functions with Aspose.Cells | batch replace GETPIVOTDATA across worksheets | structured reference example Aspose.Cells
// Developer Intent: Automatically substitute all GETPIVOTDATA formulas in a workbook with modern structured‑reference expressions.
// Use Cases: Modernize legacy reports that rely on GETPIVOTDATA before distribution. | Process large collections of workbooks to ensure compatibility with newer Excel versions. | Embed formula migration into an automated data‑export pipeline.
// AI Prompts: Write C# code using Aspose.Cells that finds every GETPIVOTDATA formula in a workbook and replaces it with a user‑defined structured reference string. | Show how to log the address, original formula, and replacement for each cell while performing the conversion. | Demonstrate parameterizing the replacement so the target table name and column are extracted from the original GETPIVOTDATA arguments.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Loads an Excel workbook, scans every worksheet and cell, detects formulas that start with GETPIVOTDATA, substitutes each with a structured‑reference placeholder (e.g., =Table1[Column1]) via Aspose.Cells, and saves the modified file.
    public class ReplaceGetPivotData
    {
        public static void Run()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            try
            {
                // Load the workbook that contains GETPIVOTDATA formulas
                Workbook workbook = new Workbook(inputPath);

                // Iterate through all worksheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Iterate through all used cells in the worksheet
                    foreach (Cell cell in sheet.Cells)
                    {
                        // Process only cells that contain a formula
                        if (cell.IsFormula)
                        {
                            string formula = cell.Formula;

                            // Check if the formula uses the deprecated GETPIVOTDATA function
                            if (formula.StartsWith("GETPIVOTDATA", StringComparison.OrdinalIgnoreCase))
                            {
                                // Placeholder structured reference replacement
                                string placeholderStructuredRef = "=Table1[Column1]";

                                // Replace the old formula with the new structured reference
                                cell.PutValue(placeholderStructuredRef);
                            }
                        }
                    }
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            ReplaceGetPivotData.Run();
        }
    }
}
