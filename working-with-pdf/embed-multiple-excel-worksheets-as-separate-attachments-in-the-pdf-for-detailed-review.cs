// Title: Create individual PDF files for each Excel worksheet and a combined PDF using Aspose.Cells in C#
// AI Prompts: Write C# code that loads an Excel workbook with Aspose.Cells, loops through all worksheets, and saves each worksheet as a separate PDF named after the sheet. | Add logic to the same program to export the whole workbook as a single combined PDF after generating the per‑sheet PDFs.
// Common Searches: Aspose.Cells C# export each worksheet to a separate PDF file | How to save an entire Excel workbook as one PDF and also get per‑sheet PDFs using Aspose.Cells | C# generate PDF files for every sheet in an Excel workbook with Aspose.Cells | Create temporary workbook for single‑sheet PDF conversion Aspose.Cells | Save Excel worksheets as individual PDFs and a combined PDF in .NET
// Tags: Aspose.Cells per‑sheet PDF export | C# generate individual worksheet PDFs | combined workbook PDF generation Aspose.Cells | temporary workbook single sheet conversion | save Excel as PDF with Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The program verifies the input Excel file, loads it with Aspose.Cells, iterates through each worksheet, creates a temporary workbook containing only that sheet, removes the default empty sheet, saves the temporary workbook as a PDF named after the sheet, and finally saves the original workbook as a single combined PDF, handling errors throughout.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input Excel file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the source Excel workbook
            Workbook workbook = new Workbook(inputPath);

            // Generate a separate PDF for each worksheet (optional demonstration)
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                try
                {
                    // Create a temporary workbook containing only the current worksheet
                    Workbook tempWb = new Workbook();
                    tempWb.Worksheets.AddCopy(sheet.Name);

                    // Remove the default empty sheet that exists in a new workbook
                    if (tempWb.Worksheets.Count > 1)
                    {
                        tempWb.Worksheets.RemoveAt(0);
                    }

                    // Save the temporary workbook as a PDF file
                    string sheetPdfPath = $"{sheet.Name}.pdf";
                    tempWb.Save(sheetPdfPath, SaveFormat.Pdf);
                    Console.WriteLine($"Generated PDF for sheet \"{sheet.Name}\": {sheetPdfPath}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to generate PDF for sheet \"{sheet.Name}\": {ex.Message}");
                }
            }

            // Save the original workbook as a single PDF
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"Combined PDF generated successfully: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
