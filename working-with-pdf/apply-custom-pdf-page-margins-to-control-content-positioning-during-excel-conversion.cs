// Title: Set custom PDF page margins for Excel worksheets when converting to PDF with Aspose.Cells for .NET
// Common Searches: how to define left and right margins for PDF output using Aspose.Cells C# | Aspose.Cells C# set page margins before saving workbook as PDF | custom PDF page margin settings in Excel to PDF conversion with Aspose.Cells | adjust PDF margins in Aspose.Cells .NET when exporting a worksheet | C# code to set top, bottom, left, right margins for PDF generated from Excel
// Tags: Aspose.Cells custom PDF margins | Excel page setup margins C# | PdfSaveOptions margin properties | FitToPagesWide PDF export Aspose.Cells | worksheet page setup for PDF conversion

using System;
using Aspose.Cells;

namespace CustomMarginsPdfExample
{
    // The example creates a Workbook, accesses the first worksheet's PageSetup, assigns left, right, top, and bottom margins in points (0.5" left/right, 0.75" top/bottom), configures FitToPagesWide, and saves the workbook as a PDF using PdfSaveOptions, demonstrating how to control PDF content positioning during Excel‑to‑PDF conversion with Aspose.Cells for .NET.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook (or load an existing one)
                Workbook workbook = new Workbook();

                // Set custom page margins for the first worksheet
                Worksheet sheet = workbook.Worksheets[0];
                PageSetup pageSetup = sheet.PageSetup;

                // Margins are specified in points (1 point = 1/72 inch)
                // Example: 0.5 inch left/right, 0.75 inch top/bottom
                pageSetup.LeftMargin = 0.5 * 72;    // 0.5 inch
                pageSetup.RightMargin = 0.5 * 72;   // 0.5 inch
                pageSetup.TopMargin = 0.75 * 72;    // 0.75 inch
                pageSetup.BottomMargin = 0.75 * 72; // 0.75 inch

                // Optional: set other page setup properties if needed
                // Orientation property removed to avoid compatibility issues
                pageSetup.FitToPagesWide = 1;
                pageSetup.FitToPagesTall = 0;

                // Convert the workbook to PDF using the defined margins
                PdfSaveOptions pdfOptions = new PdfSaveOptions();

                // Save the workbook as a PDF file
                string outputPath = "OutputWithCustomMargins.pdf";
                workbook.Save(outputPath, pdfOptions);

                Console.WriteLine($"PDF saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
