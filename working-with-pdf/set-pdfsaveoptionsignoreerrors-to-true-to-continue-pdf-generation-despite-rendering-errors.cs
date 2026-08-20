// Title: Aspose.Cells C# – Use PdfSaveOptions.IgnoreError to Continue PDF Export Despite Rendering Issues
// Description: Shows how to set PdfSaveOptions.IgnoreError = true in Aspose.Cells for .NET so PDF generation proceeds even when rendering errors occur, using a simple workbook example.
// Keywords: Aspose.Cells | PdfSaveOptions | IgnoreError | C# | .NET | PDF export | error handling | rendering errors | continue PDF generation | batch conversion | unsupported charts
// Common Searches: Aspose.Cells ignore rendering errors PDF | PdfSaveOptions.IgnoreError C# example | continue PDF export when cells cause errors | skip errors during PDF conversion Aspose.Cells | prevent PDF generation failure Aspose.Cells
// Developer Intent: Enable PdfSaveOptions.IgnoreError so the workbook can be saved as PDF even if rendering errors are encountered.
// Use Cases: Convert workbooks containing complex or unsupported charts without aborting the operation. | Batch process large numbers of spreadsheets where occasional formatting issues may trigger rendering errors. | Generate automated reports that include dynamic content which might cause intermittent rendering exceptions. | Integrate PDF export into a web service that must stay responsive despite data anomalies.
// AI Prompts: Write C# code that saves an Aspose.Cells workbook to PDF with PdfSaveOptions.IgnoreError enabled and logs any suppressed rendering errors. | Show how to combine PdfSaveOptions.IgnoreError with page orientation, compression, and image quality settings in Aspose.Cells. | Explain the visual impact of setting PdfSaveOptions.IgnoreError to true on the resulting PDF and how to programmatically detect suppressed errors.

using System;
using Aspose.Cells;

namespace AsposeCellsPdfIgnoreErrorDemo
{
    // Shows how to set PdfSaveOptions.IgnoreError = true in Aspose.Cells for .NET so PDF generation proceeds even when rendering errors occur, using a simple workbook example.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (using the standard creation rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet and add some sample data
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample data");
            sheet.Cells["A2"].PutValue(123);
            sheet.Cells["A3"].PutValue(DateTime.Now);

            // Create PDF save options (using the provided constructor rule)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Set IgnoreError to true so that rendering errors are hidden and PDF generation continues
            pdfOptions.IgnoreError = true;

            // Save the workbook as PDF with the specified options (using the standard save rule)
            workbook.Save("OutputWithIgnoreError.pdf", pdfOptions);

            Console.WriteLine("PDF saved successfully with IgnoreError enabled.");
        }
    }
}
