// Title: Export a selected worksheet range to PDF with custom margins and landscape orientation using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# export specific range to PDF with custom margins | Set page orientation to landscape when saving an Excel range as PDF using Aspose.Cells | Define print area and margins before converting a worksheet to PDF in .NET | How does PdfSaveOptions OnePagePerSheet affect range export in Aspose.Cells | C# code to save an Excel print area as PDF with 0.75 inch margins
// Tags: range to PDF export Aspose.Cells | custom page margins Aspose.Cells | landscape orientation PDF Aspose.Cells | print area PDF conversion Aspose.Cells | PdfSaveOptions OnePagePerSheet false

using System;
using Aspose.Cells;

// The example creates a workbook, sets the print area to A1:C3, applies 0.75‑inch margins on all sides, changes the page orientation to landscape, configures PdfSaveOptions to keep the range on a single page, and saves the defined range as a PDF file.
class ExportRangeToPdf
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data (replace with your own data as needed)
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Quantity");
            sheet.Cells["C1"].PutValue("Price");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(50);
            sheet.Cells["C2"].PutValue(0.5);
            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["B3"].PutValue(30);
            sheet.Cells["C3"].PutValue(0.3);

            // Define the range that should be exported as PDF
            string exportRange = "A1:C3";

            // Set the print area to the defined range
            sheet.PageSetup.PrintArea = exportRange;

            // Set custom margins (values are in points; 1 inch = 72 points)
            double marginInInches = 0.75;
            double marginInPoints = marginInInches * 72;
            sheet.PageSetup.LeftMargin   = marginInPoints;
            sheet.PageSetup.RightMargin  = marginInPoints;
            sheet.PageSetup.TopMargin    = marginInPoints;
            sheet.PageSetup.BottomMargin = marginInPoints;

            // Set page orientation (Portrait or Landscape)
            sheet.PageSetup.Orientation = PageOrientationType.Landscape;

            // Configure PDF save options (optional)
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Ensure only the defined print area is saved, not the whole sheet
                OnePagePerSheet = false
            };

            // Export the defined range to a PDF file with the specified settings
            workbook.Save("ExportedRange.pdf", pdfOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
