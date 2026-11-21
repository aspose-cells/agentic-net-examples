// Title: How to export large numeric values to HTML without scientific notation using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# export large integer to HTML without scientific notation | How to keep full numeric value in HTML output using Aspose.Cells .NET | Set custom number format for a cell before saving as HTML in Aspose.Cells | Prevent exponential format for big numbers in Aspose.Cells HTMLSaveOptions | C# Aspose.Cells HTML export retain decimal representation for large numbers
// Tags: HTML export custom number format Aspose.Cells | suppress scientific notation Aspose.Cells | large integer HTML output .NET | cell style custom format Aspose.Cells | HtmlSaveOptions number format control

using System;
using Aspose.Cells;

namespace LargeNumberExportExample
{
    // The example creates a workbook, writes a very large double to cell A1, applies a custom style with the format "0" to force plain decimal display, and saves the workbook as HTML using HtmlSaveOptions, ensuring the exported HTML shows the full number without scientific (exponential) notation.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Insert a large numeric value that would normally be shown in exponential notation
                double largeNumber = 123456789012345.0;
                sheet.Cells["A1"].PutValue(largeNumber);

                // Define a custom number format that forces plain decimal representation (no scientific notation)
                Style style = workbook.CreateStyle();
                style.Custom = "0"; // Shows the full integer without exponent
                sheet.Cells["A1"].SetStyle(style);

                // Prepare HTML save options
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
                // The default behavior exports the cell's formatted value, which preserves the custom format.

                // Save the workbook as HTML
                string outputPath = "LargeNumberExport.html";
                workbook.Save(outputPath, htmlOptions);

                Console.WriteLine($"Workbook successfully saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while exporting the workbook:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
