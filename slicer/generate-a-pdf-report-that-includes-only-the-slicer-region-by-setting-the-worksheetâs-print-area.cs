// Title: Generate a PDF containing only a slicer region by setting the worksheet PrintArea with Aspose.Cells for .NET (C#)
// AI Prompts: Write a C# routine that assigns the slicer's cell range to the worksheet PrintArea and saves the workbook as a PDF using Aspose.Cells. | Enhance the code to automatically locate the slicer on the sheet, set its bounds as the PrintArea, and then export to PDF. | Add validation logic that checks whether the specified PrintArea exists and contains a slicer before performing the PDF conversion.
// Common Searches: Aspose.Cells C# set print area to slicer range before PDF export | How to export only the slicer region to PDF using Aspose.Cells .NET | C# generate PDF with slicer only using Aspose.Cells | Configure worksheet PrintArea for slicer and save as PDF Aspose.Cells example | Aspose.Cells PDFSaveOptions print area slicer example
// Tags: set worksheet printarea for slicer pdf export | Aspose.Cells PDFSaveOptions slicer region | C# export slicer area to PDF | configure print area to slicer range Aspose.Cells | generate pdf from workbook with slicer only

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering; // Optional: for ImageOrPrintOptions if needed in other scenarios

namespace SlicerPdfReport
{
    // Creates SlicerReport.pdf that includes only the cells covering the slicer (e.g., C5:F10) by setting the worksheet's PrintArea and saving the workbook with PdfSaveOptions.
    public class GenerateReport
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // ------------------------------------------------------------
                // Populate some data (optional – just to have content in the sheet)
                // ------------------------------------------------------------
                sheet.Cells["A1"].PutValue("Product");
                sheet.Cells["B1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("Apple");
                sheet.Cells["B2"].PutValue("Fruit");
                sheet.Cells["A3"].PutValue("Carrot");
                sheet.Cells["B3"].PutValue("Vegetable");
                // ... add more rows as needed ...

                // ------------------------------------------------------------
                // Assume a slicer has been added to the worksheet covering cells C5:F10.
                // Set the print area to that exact range so that only the slicer region
                // will be included in the exported PDF.
                // ------------------------------------------------------------
                sheet.PageSetup.PrintArea = "C5:F10";

                // ------------------------------------------------------------
                // Configure PDF save options (default options are sufficient for this case)
                // ------------------------------------------------------------
                PdfSaveOptions pdfOptions = new PdfSaveOptions();

                // ------------------------------------------------------------
                // Save the workbook as PDF. Because the PrintArea is set, only the
                // slicer region (C5:F10) will appear in the resulting PDF file.
                // ------------------------------------------------------------
                workbook.Save("SlicerReport.pdf", pdfOptions);

                Console.WriteLine("PDF report with slicer region created successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred while generating the PDF report: {ex.Message}");
            }
        }
    }

    // Entry point for the console application
    public class Program
    {
        public static void Main(string[] args)
        {
            GenerateReport.Run();
        }
    }
}
