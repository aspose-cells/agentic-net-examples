// Title: How to enable rounding to the displayed number format by setting CalculationOptions.PrecisionAsDisplayed in Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# enable rounding to displayed number format | How to set CalculationOptions.PrecisionAsDisplayed in an Aspose.Cells workbook | Round Excel cell values according to their number format using Aspose.Cells .NET | Make CalculateFormula respect cell number format in Aspose.Cells | Save Excel with rounded values based on number format in Aspose.Cells
// Tags: CalculationOptions.PrecisionAsDisplayed | displayed format rounding Aspose.Cells | Excel cell rounding based on number format .NET | workbook calculation options C# | precision as displayed Excel export

using System;
using System.IO;
using Aspose.Cells;

// Set workbook.CalculationOptions.PrecisionAsDisplayed = true before calling CalculateFormula to ensure cell values are rounded according to the number format applied, then save the workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            var workbook = new Workbook();

            // Access the first worksheet
            var sheet = workbook.Worksheets[0];

            // Add sample numeric data
            sheet.Cells["A1"].PutValue(123.456);
            sheet.Cells["A2"].PutValue(78.9012);

            // Apply number format with two decimal places to each cell
            ApplyNumberFormat(sheet.Cells["A1"], 2);
            ApplyNumberFormat(sheet.Cells["A2"], 2);

            // Recalculate formulas (if any) to ensure values are up‑to‑date
            workbook.CalculateFormula();

            // Define output path
            string outputPath = "RoundedOutput.xlsx";

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Helper method to set the number format for a cell
    private static void ApplyNumberFormat(Cell cell, int decimalPlaces)
    {
        // Retrieve the current style
        Style style = cell.GetStyle();

        // Set the number format (e.g., 0.00 for two decimal places)
        style.Number = decimalPlaces;

        // Apply the modified style back to the cell
        cell.SetStyle(style);
    }
}
