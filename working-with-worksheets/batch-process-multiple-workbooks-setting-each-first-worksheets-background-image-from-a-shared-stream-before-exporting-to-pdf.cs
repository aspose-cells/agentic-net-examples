// Title: Batch convert Excel workbooks to PDF with a shared first‑worksheet background image using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads several .xlsx files, sets a PNG as the background of the first sheet in each workbook, and converts them to PDF with Aspose.Cells. | Refactor the example to read the background image once into a byte array and reuse that byte array for all workbook conversions, reducing file I/O. | Add robust error handling that skips missing Excel files, logs each conversion result, and continues processing the remaining workbooks.
// Common Searches: Aspose.Cells set identical background picture on first worksheet for multiple Excel files before PDF conversion | C# batch convert Excel to PDF while reusing a single image stream for worksheet backgrounds | How to skip non‑existent Excel files during bulk PDF export with Aspose.Cells | Optimize image loading when applying background to many workbooks in Aspose.Cells
// Tags: batch workbook to PDF Aspose.Cells | worksheet background from byte array | reuse single image stream for multiple workbooks | first sheet background image C# | handle missing Excel files during conversion

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The program iterates over a list of Excel file paths, loads a PNG background image once, assigns it to the first worksheet of each workbook, and saves each workbook as a PDF, while gracefully handling missing files and reporting conversion outcomes.
class Program
{
    static void Main()
    {
        // List of workbook files to process
        var workbookPaths = new List<string>
        {
            "input1.xlsx",
            "input2.xlsx",
            // add additional workbook file paths here
        };

        // Path to the background image file
        const string backgroundImagePath = "background.png";

        try
        {
            // Verify that the background image exists
            if (!File.Exists(backgroundImagePath))
                throw new FileNotFoundException("Background image file not found.", backgroundImagePath);

            // Load the background image into a byte array once for reuse
            byte[] imageBytes = File.ReadAllBytes(backgroundImagePath);

            foreach (string wbPath in workbookPaths)
            {
                // Verify that the workbook file exists
                if (!File.Exists(wbPath))
                {
                    Console.WriteLine($"Workbook file not found: {wbPath}");
                    continue;
                }

                try
                {
                    // Load the workbook
                    var workbook = new Workbook(wbPath);

                    // Access the first worksheet
                    Worksheet firstSheet = workbook.Worksheets[0];

                    // Apply the background image (expects a byte array)
                    firstSheet.BackgroundImage = imageBytes;

                    // Determine output PDF path
                    string pdfPath = Path.ChangeExtension(wbPath, ".pdf");

                    // Save as PDF
                    workbook.Save(pdfPath, SaveFormat.Pdf);
                    Console.WriteLine($"Converted '{wbPath}' to PDF successfully.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing workbook '{wbPath}': {ex.Message}");
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Fatal error: {ex.Message}");
        }
    }
}
