// Title: Replace CONCATENATE formulas with the '&' operator in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# replace CONCATENATE function with '&' operator in Excel formulas | convert legacy CONCATENATE formulas to '&' using Aspose.Cells .NET | update Excel workbook formulas from CONCATENATE to '&' with Aspose.Cells | iterate cells and modify CONCATENATE formulas in C# Aspose.Cells | replace CONCATENATE with modern concatenation operator in .xlsx using Aspose.Cells
// Tags: Aspose.Cells formula conversion CONCATENATE to ampersand | C# batch Excel formula update Aspose.Cells | programmatic Excel concatenation operator migration | Excel .xlsx formula rewrite using Aspose.Cells .NET | automated legacy formula replacement in C#

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an input.xlsx workbook with Aspose.Cells, scans every worksheet and cell, detects formulas that use the CONCATENATE function, extracts their arguments, rebuilds each formula using the '&' concatenation operator, writes the new formula back to the cell, and saves the modified workbook as output.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Iterate through all worksheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    Cells cells = sheet.Cells;

                    // Iterate through all used cells
                    foreach (Cell cell in cells)
                    {
                        // Process only formula cells containing CONCATENATE
                        if (cell.IsFormula && cell.Formula.IndexOf("CONCATENATE", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            string originalFormula = cell.Formula; // e.g., "=CONCATENATE(A1, \" \", B1)"

                            int funcStart = originalFormula.IndexOf("CONCATENATE", StringComparison.OrdinalIgnoreCase);
                            int openParen = originalFormula.IndexOf('(', funcStart);
                            int closeParen = originalFormula.LastIndexOf(')');

                            if (openParen > 0 && closeParen > openParen)
                            {
                                // Extract arguments inside CONCATENATE(...)
                                string formulaArgs = originalFormula.Substring(openParen + 1, closeParen - openParen - 1);

                                // Split arguments by commas (simple split; assumes no commas inside quoted strings)
                                string[] parts = formulaArgs.Split(',');

                                // Trim whitespace from each part
                                for (int i = 0; i < parts.Length; i++)
                                {
                                    parts[i] = parts[i].Trim();
                                }

                                // Build the new formula using the modern CONCAT operator (&)
                                string newFormula = "=" + string.Join("&", parts);

                                // Assign the transformed formula back to the cell
                                cell.Formula = newFormula;
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
                Console.WriteLine("An error occurred while processing the workbook:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
