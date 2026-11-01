// Title: How to replace CONCATENATE formulas with the CONCAT operator in an Excel workbook using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# replace CONCATENATE function with CONCAT in all worksheets | Update Excel formulas from CONCATENATE to CONCAT using Aspose.Cells .NET | Programmatically change CONCATENATE to CONCAT operator in a workbook with Aspose.Cells | C# code to convert legacy CONCATENATE formulas to modern CONCAT in Excel files
// Tags: replace CONCATENATE with CONCAT using Aspose.Cells | bulk formula update Excel .NET | regex formula transformation Aspose.Cells | recalculate workbook after formula change Aspose.Cells | save modified Excel workbook Aspose.Cells

using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;

// The C# example loads an existing workbook (creating a blank one if missing), scans every worksheet and cell for formulas containing the legacy CONCATENATE function, replaces it with the modern CONCAT operator using a case‑insensitive regular expression, recalculates the workbook, and saves the updated file to a new location with Aspose.Cells.
class ReplaceConcatenateWithConcat
{
    static void Main()
    {
        // Define input and output file paths
        string inputPath = @"C:\Docs\InputWorkbook.xlsx";
        string outputPath = @"C:\Docs\OutputWorkbook.xlsx";

        try
        {
            // Verify that the input file exists; create a blank workbook if it does not
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                Workbook emptyWb = new Workbook();
                // Ensure the directory for the input file exists
                string inputDir = Path.GetDirectoryName(inputPath);
                if (!string.IsNullOrEmpty(inputDir) && !Directory.Exists(inputDir))
                {
                    Directory.CreateDirectory(inputDir);
                }
                emptyWb.Save(inputPath);
                Console.WriteLine($"Created empty workbook at {inputPath}");
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Regular expression to replace CONCATENATE( with CONCAT(
            Regex concatRegex = new Regex(@"CONCATENATE\s*\(", RegexOptions.IgnoreCase);

            // Iterate through all worksheets and cells containing formulas
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (Cell cell in sheet.Cells)
                {
                    if (cell.IsFormula)
                    {
                        string originalFormula = cell.Formula;
                        if (originalFormula.IndexOf("CONCATENATE", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            // Replace CONCATENATE with CONCAT
                            string updatedFormula = concatRegex.Replace(originalFormula, "CONCAT(");
                            cell.Formula = updatedFormula;
                        }
                    }
                }
            }

            // Recalculate all formulas after modifications
            workbook.CalculateFormula();

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
