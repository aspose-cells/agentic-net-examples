// Title: How to enable rounding to the displayed cell format in Aspose.Cells for .NET using CalculationOptions.PrecisionAsDisplayed
// Common Searches: Aspose.Cells set PrecisionAsDisplayed to true for rounding numbers | C# round Excel cell values according to display format with Aspose.Cells | Enable displayed precision when saving workbook using Aspose.Cells .NET | CalculationOptions.PrecisionAsDisplayed usage example in C# | Round numbers to two decimal places based on cell style in Aspose.Cells
// Tags: CalculationOptions.PrecisionAsDisplayed property | rounding numbers to displayed format Aspose.Cells | apply built‑in number format with rounding .NET | Excel workbook calculation options Aspose.Cells | C# set precision as displayed for Excel export

using System;
using Aspose.Cells;

// The example creates a new Workbook, writes numeric values to cells A1 and A2, applies a built‑in two‑decimal‑place number format, enables CalculationOptions.PrecisionAsDisplayed so values are rounded to the displayed format, recalculates any formulas, and saves the result as RoundedOutput.xlsx.
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

            // Sample data
            sheet.Cells["A1"].PutValue(123.4567);
            sheet.Cells["A2"].PutValue(0.123456);

            // Set display format to two decimal places (built‑in format "0.00")
            Style style = workbook.CreateStyle();
            style.Number = 2; // Built‑in format index for two decimal places
            sheet.Cells["A1"].SetStyle(style);
            sheet.Cells["A2"].SetStyle(style);

            // Recalculate formulas (if any)
            workbook.CalculateFormula();

            // Save the workbook
            string outputPath = "RoundedOutput.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
