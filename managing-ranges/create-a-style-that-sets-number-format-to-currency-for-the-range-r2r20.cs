using System;
using System.IO;
using Aspose.Cells;

class CurrencyStyleExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Create a style and set its number format to a built‑in currency format (value 5)
            Style currencyStyle = workbook.CreateStyle();
            currencyStyle.Number = 5; // "$#,##0_);($#,##0)" for en_US region

            // Define the target range R2:R20 (use fully qualified type to avoid ambiguity with System.Range)
            Aspose.Cells.Range targetRange = worksheet.Cells.CreateRange("R2", "R20");

            // Apply the currency style to the entire range
            targetRange.SetStyle(currencyStyle);

            // Ensure the output directory exists
            string outputPath = "CurrencyStyle.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}