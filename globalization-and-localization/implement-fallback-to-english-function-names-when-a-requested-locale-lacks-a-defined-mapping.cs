// Title: How to use English Excel formulas as a fallback when a locale has no mapping in Aspose.Cells for C#
// Common Searches: Aspose.Cells C# use English formulas when locale mapping is missing | fallback to English function names in Aspose.Cells for unsupported culture | calculate workbook formulas with undefined locale in Aspose.Cells | set workbook CultureInfo to de-DE but keep English formulas Aspose.Cells | handle missing locale mapping for Excel functions in Aspose.Cells C#
// Tags: Aspose.Cells default English formulas | Aspose.Cells locale mapping absence handling | Workbook CultureInfo configuration Aspose.Cells | Excel formula localization strategy Aspose.Cells | C# fallback for undefined culture in Aspose.Cells

using System;
using System.Globalization;
using System.IO;
using Aspose.Cells;

// The example creates a workbook, sets its CultureInfo to German (de-DE), writes English formulas (SUM and AVERAGE) directly into cells, calculates the formulas, displays the results, and saves the file as FallbackExample.xlsx, demonstrating how English function names act as a fallback when a locale lacks a specific mapping.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Set the workbook locale to German (de-DE)
            workbook.Settings.CultureInfo = new CultureInfo("de-DE");

            // Populate some data
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].PutValue(20);
            sheet.Cells["A3"].PutValue(30);

            // Use English function names (fallback behavior is not needed without custom mapping)
            sheet.Cells["B1"].Formula = "SUM(A1:A3)";
            sheet.Cells["B2"].Formula = "AVERAGE(A1:A3)";

            // Calculate all formulas
            workbook.CalculateFormula();

            // Output the results
            Console.WriteLine("SUM result (should be 60): " + sheet.Cells["B1"].Value);
            Console.WriteLine("AVERAGE result (should be 20): " + sheet.Cells["B2"].Value);

            // Define output file path
            string outputPath = "FallbackExample.xlsx";

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
