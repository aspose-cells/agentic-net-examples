// Title: How to unhide columns N‑P, set a 70‑point column width, and export an Excel worksheet to PDF using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# unhide columns N to P and set column width in points | Set column width to 70 points and export worksheet as PDF with Aspose.Cells | Convert Excel to PDF after adjusting column visibility using Aspose.Cells .NET | C# code to unhide specific columns and define column width before PDF conversion | How to change column width in pixels for Aspose.Cells before saving as PDF
// Tags: unhide specific column range Aspose.Cells | set column width in points Aspose.Cells | convert Excel worksheet to PDF Aspose.Cells | column width pixel calculation Aspose.Cells | adjust column visibility before PDF export Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example loads an existing Excel workbook, unhides columns N through P on the first worksheet, sets each column's width to 70 points (converted to pixels), and then saves the workbook as a PDF, with error handling for missing files, loading issues, column processing, and PDF export.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.pdf";

        // Verify that the input file exists before attempting to load it
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        Workbook workbook;
        try
        {
            // Load the workbook from the existing file
            workbook = new Workbook(inputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error loading workbook: {ex.Message}");
            return;
        }

        // Get the first worksheet (adjust index if needed)
        Worksheet sheet = workbook.Worksheets[0];

        // Columns N (index 13) through P (index 15)
        for (int col = 13; col <= 15; col++)
        {
            try
            {
                // Unhide the column using the Columns collection
                sheet.Cells.Columns[col].IsHidden = false;

                // Set column width to 70 points.
                // Convert points to pixels: 1 point = 1/72 inch, 96 DPI => pixels = points * 96 / 72
                int pixelWidth = (int)Math.Ceiling(70.0 * 96.0 / 72.0);
                sheet.Cells.SetColumnWidthPixel(col, pixelWidth);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing column {col}: {ex.Message}");
            }
        }

        try
        {
            // Save the workbook as a PDF file
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"Workbook successfully saved as \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error saving PDF: {ex.Message}");
        }
    }
}
