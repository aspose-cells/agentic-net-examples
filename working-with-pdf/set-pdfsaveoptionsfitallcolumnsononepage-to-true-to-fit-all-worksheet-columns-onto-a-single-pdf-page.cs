// Title: How to export an Excel worksheet to a single-page PDF by enabling PdfSaveOptions.FitAllColumnsOnOnePage in C# with Aspose.Cells
// AI Prompts: Generate C# code that creates or loads a Workbook, sets PdfSaveOptions.FitAllColumnsOnOnePage = true, and saves the workbook as a PDF where all columns fit on one page using Aspose.Cells. | Update the provided example to replace the PageSetup FitToPagesWide/Tall settings with PdfSaveOptions.FitAllColumnsOnOnePage, ensuring the PDF output contains a single page per sheet. | Write a reusable C# method that accepts a Workbook and an output path, configures PdfSaveOptions to fit all columns on one PDF page, and writes the PDF with Aspose.Cells.
// Common Searches: Aspose.Cells C# PdfSaveOptions FitAllColumnsOnOnePage usage example | export Excel to PDF with all columns on one page using Aspose.Cells | set FitAllColumnsOnOnePage true for PDF conversion in Aspose.Cells C# | single-page PDF from worksheet columns Aspose.Cells tutorial
// Tags: Aspose.Cells PdfSaveOptions column fit on PDF | single-page PDF export Aspose.Cells C# | fit all worksheet columns PDF Aspose.Cells | C# Aspose.Cells PDF conversion options | one page per sheet PDF save Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample creates a Workbook, adds sample data, configures PdfSaveOptions with FitAllColumnsOnOnePage set to true and OnePagePerSheet enabled, then saves the workbook as a PDF where every worksheet's columns are scaled to fit on a single page.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Populate sample data
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Header1");
            sheet.Cells["B1"].PutValue("Header2");
            sheet.Cells["A2"].PutValue("Data1");
            sheet.Cells["B2"].PutValue("Data2");

            // Configure page setup to fit all columns on one page
            sheet.PageSetup.FitToPagesWide = 1;   // fit columns to one page width
            sheet.PageSetup.FitToPagesTall = 0;   // unlimited rows (no vertical scaling)

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                OnePagePerSheet = true
            };

            // Save the workbook as PDF using the configured options
            string outputPath = "output.pdf";
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
