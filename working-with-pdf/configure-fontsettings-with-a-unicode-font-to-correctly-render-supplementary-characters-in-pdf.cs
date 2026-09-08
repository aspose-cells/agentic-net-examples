// Title: Configure FontSettings with a Unicode font to render supplementary characters (e.g., emoji) in a PDF using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that sets a Unicode‑compatible font on a cell and saves the workbook as a PDF with Aspose.Cells. | Show how to use FontSettings in Aspose.Cells to ensure supplementary Unicode characters are displayed correctly in PDF output. | Provide a complete Aspose.Cells example that inserts an emoji, applies Arial Unicode MS, and exports the workbook to PDF.
// Common Searches: Aspose.Cells C# how to display emoji in PDF export | set Unicode font for supplementary characters when saving workbook as PDF with Aspose.Cells | FontSettings configuration for Unicode support in Aspose.Cells PDF generation .NET
// Tags: FontSettings Unicode font Aspose.Cells PDF | supplementary character rendering Aspose.Cells C# | emoji support in PDF export Aspose.Cells | Arial Unicode MS font for PDF generation Aspose.Cells | cell style font configuration Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // // Example creates a workbook, places an emoji (😀) in cell A1, applies the Arial Unicode MS font via FontSettings to support supplementary Unicode characters, and saves the workbook as a PDF, ensuring correct rendering of the emoji.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                var workbook = new Workbook();

                // Access the first worksheet
                var worksheet = workbook.Worksheets[0];

                // Put a supplementary Unicode character (e.g., 😀) into a cell
                var cell = worksheet.Cells["A1"];
                cell.PutValue("😀");

                // Apply a font that supports supplementary characters to the cell
                var style = cell.GetStyle();
                style.Font.Name = "Arial Unicode MS";
                cell.SetStyle(style);

                // Define output file path
                string outputPath = "SupplementaryCharacters.pdf";

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook as PDF; the supplementary character will be rendered correctly
                workbook.Save(outputPath, SaveFormat.Pdf);

                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
