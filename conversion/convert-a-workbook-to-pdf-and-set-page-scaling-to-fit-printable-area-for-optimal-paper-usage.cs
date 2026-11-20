// Title: Convert an Excel workbook to PDF with fit‑to‑page scaling in C# using Aspose.Cells
// Common Searches: c# aspnet convert excel file to pdf with fit to page width using Aspose.Cells | set worksheet page setup to fit one page wide before exporting to PDF with Aspose.Cells | Aspose.Cells PDF export fit to printable area example in C# | how to use FitToPagesWide and FitToPagesTall in Aspose.Cells when saving as PDF | generate PDF from Excel with optimal paper usage Aspose.Cells C#
// Tags: Aspose.Cells PDF export with page scaling | FitToPagesWide worksheet scaling Aspose.Cells | Excel to PDF conversion C# Aspose.Cells | page setup fit-to-page before PDF save | optimize PDF paper usage Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The sample loads an existing Excel workbook (or creates a simple one if missing), configures each worksheet's PageSetup to fit the content to one page wide while allowing the height to adjust automatically, and then saves the workbook as a PDF using Aspose.Cells, ensuring optimal paper usage.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.pdf";

        Workbook workbook = null;

        // Ensure the input file exists; create a simple workbook if it does not.
        if (File.Exists(inputPath))
        {
            try
            {
                workbook = new Workbook(inputPath);
            }
            catch (CellsException ex)
            {
                Console.WriteLine($"Failed to load workbook '{inputPath}': {ex.Message}");
                return;
            }
        }
        else
        {
            try
            {
                workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "Sheet1";
                sheet.Cells["A1"].PutValue("Sample data");
                workbook.Save(inputPath);
                Console.WriteLine($"Created sample workbook '{inputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error creating sample workbook: {ex.Message}");
                return;
            }
        }

        // Apply page setup to fit each worksheet to a single page width.
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            PageSetup pageSetup = sheet.PageSetup;
            pageSetup.FitToPagesWide = 1;   // Fit to one page wide
            pageSetup.FitToPagesTall = 0;   // Auto-adjust height
        }

        // Save the workbook as PDF.
        try
        {
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"Workbook saved as PDF to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error saving PDF: {ex.Message}");
        }
    }
}
