// Title: Save an Excel workbook with German culture-specific number formatting using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that creates a new Workbook, inserts a numeric value, applies a two‑decimal style, assigns the workbook's CultureInfo to "de-DE", and writes the file to disk. | Show how to configure Aspose.Cells to produce locale‑aware numeric output (e.g., German decimal separators) before saving an Excel document.
// Common Searches: Aspose.Cells C# how to apply German number formatting to a workbook | Set workbook CultureInfo to de-DE for Excel export using Aspose.Cells | Locale‑specific numeric style in Aspose.Cells .NET example | Saving Excel file with German decimal separators via Aspose.Cells | Configure workbook settings for culture‑aware formatting in C#
// Tags: Aspose.Cells workbook cultureinfo configuration | C# German locale number style in Excel | Excel export with culture‑specific numeric format | Aspose.Cells set de-DE number formatting | save workbook with locale‑aware formatting

using System;
using System.Globalization;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The program creates a new Workbook, writes the value 12345.67 to cell A1, applies a two‑decimal number style, sets Workbook.Settings.CultureInfo to the German locale (de-DE) so that separators follow German conventions, ensures the output directory exists, and saves the file as CultureSpecificNumberFormats.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Put a numeric value into a cell
                sheet.Cells["A1"].PutValue(12345.67);

                // Apply a number format (two decimal places)
                Style style = sheet.Cells["A1"].GetStyle();
                style.Number = 10; // "#,##0.00"
                sheet.Cells["A1"].SetStyle(style);

                // Set the workbook culture to German (Germany) for locale‑specific formatting
                workbook.Settings.CultureInfo = new CultureInfo("de-DE");

                // Define output file path
                string outputPath = "CultureSpecificNumberFormats.xlsx";

                // Ensure the directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook; the number format will respect the specified locale
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
