// Title: Replace simple INDIRECT formulas with direct cell references in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx file with Aspose.Cells, scans all worksheets, and substitutes any INDIRECT("A1") or INDIRECT("Sheet1!B2") formulas with the equivalent direct reference. | Write a method that uses a regular expression to locate INDIRECT calls in cell formulas, rewrites them to plain references, and saves the updated workbook to a new path.
// Common Searches: how to eliminate INDIRECT function from formulas with Aspose.Cells C# | Aspose.Cells replace indirect references in bulk across worksheets | C# code to convert INDIRECT formulas to direct cell addresses in Excel files | improve Excel calculation performance by removing INDIRECT using .NET library | regex based formula cleanup Aspose.Cells example
// Tags: replace indirect formulas Aspose.Cells | direct cell reference conversion .NET | regex formula transformation Aspose.Cells | bulk worksheet formula editing C# | calculation stability improvement Excel .NET

using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;

// The example loads an Excel workbook with Aspose.Cells, iterates through each worksheet and used cells, detects formulas containing the INDIRECT function, and uses a regular expression to replace simple INDIRECT("A1") or INDIRECT("Sheet1!B2") calls with the raw cell reference. Modified formulas are written back to the cells, and the workbook is saved to a new file, improving calculation stability.
class ReplaceIndirectFormulas
{
    static void Main()
    {
        try
        {
            // Input workbook path – adjust as needed
            string inputPath = @"C:\Path\To\InputWorkbook.xlsx";

            // Verify the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Regex to capture simple INDIRECT references like INDIRECT("A1") or INDIRECT("Sheet1!B2")
            Regex indirectPattern = new Regex(@"INDIRECT\(\s*""([^""]+)""\s*\)", RegexOptions.IgnoreCase);

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;

                // Iterate over used cells only for better performance
                foreach (Cell cell in cells)
                {
                    // Process cells that contain a formula
                    if (cell.IsFormula)
                    {
                        string formula = cell.Formula;

                        // Quick check for the presence of INDIRECT
                        if (formula.IndexOf("INDIRECT", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            // Replace simple INDIRECT calls with direct references
                            string newFormula = indirectPattern.Replace(formula, match =>
                            {
                                // Extract the inner reference (e.g., A1 or Sheet1!B2)
                                string innerReference = match.Groups[1].Value;
                                // Return the direct reference without the INDIRECT wrapper
                                return innerReference;
                            });

                            // Update the cell only if a change occurred
                            if (!newFormula.Equals(formula, StringComparison.Ordinal))
                            {
                                cell.Formula = newFormula;
                            }
                        }
                    }
                }
            }

            // Output workbook path – adjust as needed
            string outputPath = @"C:\Path\To\OutputWorkbook.xlsx";

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to: {outputPath}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
