// Title: How to enable automatic formula recalculation for German‑locale formulas using Workbook.Settings in Aspose.Cells for .NET
// Common Searches: set Workbook.Settings.AutoCalculate true to recalculate German locale formulas in Aspose.Cells C# | Aspose.Cells automatically calculate formulas after using ';' as argument separator in German Excel files | C# example for enabling auto‑recalc of localized Excel formulas with Aspose.Cells | how to make Aspose.Cells recalculate formulas when using German locale separators | automatic calculation of SUM formula with semicolon separator in Aspose.Cells .NET
// Tags: Workbook.Settings AutoCalculate | German locale formula separator | automatic formula recalculation Aspose.Cells | localized Excel formulas C# | save workbook with auto calculate

using System;
using System.IO;
using Aspose.Cells;

// The example creates a workbook, fills cells A1 and A2, assigns a German‑locale SUM formula using ';' as the separator, enables automatic recalculation via Workbook.Settings.AutoCalculate, and saves the result as an XLSX file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate example data
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].PutValue(20);

            // Apply a localized formula (German locale uses ';' as argument separator)
            sheet.Cells["A3"].Formula = "=SUM(A1;A2)";

            // Force recalculation to ensure the formula result is up‑to‑date
            workbook.CalculateFormula();

            // Define output file path
            string outputPath = "output.xlsx";

            // Ensure the directory for the output file exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
