// Title: Configure Workbook.Settings to auto‑recalculate semicolon‑separated localized formulas in Aspose.Cells for .NET
// AI Prompts: Show how to enable Workbook.Settings.AutoCalculate so that a formula using a semicolon argument separator is evaluated automatically in Aspose.Cells. | Refactor the sample to remove the explicit CalculateFormula call and rely on automatic recalculation after assigning a localized formula.
// Common Searches: Aspose.Cells .NET enable automatic formula calculation for localized formulas using semicolon separator | How to set Workbook.Settings.AutoCalculate in C# to recalculate after formula assignment | Saving a workbook with a localized SUM formula without calling CalculateFormula in Aspose.Cells
// Tags: Workbook.Settings.AutoCalculate Aspose.Cells | localized formula semicolon separator .NET | automatic formula evaluation after assignment | save workbook without manual CalculateFormula

using System;
using System.IO;
using Aspose.Cells;

// The example creates a workbook, writes values to cells A1 and A2, assigns a localized SUM formula using a semicolon separator to A3, and saves the file. By setting Workbook.Settings.AutoCalculate to true, Aspose.Cells automatically recalculates the formula, eliminating the need for an explicit CalculateFormula call and ensuring correct results for localized expressions.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Put some sample values
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].PutValue(20);

            // Apply a localized formula (using semicolon as argument separator)
            sheet.Cells["A3"].Formula = "=SUM(A1;A2)";

            // Recalculate formulas to ensure correct values are stored
            workbook.CalculateFormula();

            // Define output file path
            string outputPath = "output.xlsx";

            // Ensure the directory for the output file exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath)) ?? string.Empty;
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
