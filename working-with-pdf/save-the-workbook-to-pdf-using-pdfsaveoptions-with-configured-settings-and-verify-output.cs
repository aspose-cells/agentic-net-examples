using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Author: Aspose.Cells .NET example

class PdfSaveDemo
{
    static void Main()
    {
        // Create a new workbook and add sample data
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Name = "DemoSheet";
        sheet.Cells["A1"].PutValue("Aspose.Cells PDF Save Demo");
        sheet.Cells["A2"].PutValue(DateTime.Now);
        sheet.Cells["A3"].PutValue(12345.67);
        sheet.Cells["A4"].PutValue(true);

        // Configure PDF save options
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Set PDF/A-1b compliance
            Compliance = PdfCompliance.PdfA1b,
            // Do not check workbook default font (faster rendering)
            CheckWorkbookDefaultFont = false,
            // Render each sheet on a separate page
            OnePagePerSheet = true,
            // Embed standard Windows fonts
            EmbedStandardWindowsFonts = true,
            // Set a custom producer string
            Producer = "Aspose.Cells PDF Demo"
        };

        // Define output file path
        string outputPath = "DemoOutput.pdf";

        // Save the workbook to PDF using the configured options
        workbook.Save(outputPath, pdfOptions);

        // Verify that the PDF file was created
        if (File.Exists(outputPath))
        {
            Console.WriteLine($"PDF saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        else
        {
            Console.WriteLine("Failed to save PDF.");
        }
    }
}