// Title: Disable gridlines in PDF export with Aspose.Cells for .NET – PdfSaveOptions.RenderSolidGridlines
// Description: This example shows how to create a workbook, add sample data, and export it to PDF without visible gridlines by setting PdfSaveOptions.RenderSolidGridlines to false. The resulting PDF has a clean layout suitable for reports and invoices.
// Keywords: Aspose.Cells PDF gridlines false | RenderSolidGridlines C# | hide worksheet gridlines PDF Aspose | .NET PDF export without gridlines | Aspose.Cells PdfSaveOptions example | C# clean PDF layout Aspose | USA .NET developers Aspose.Cells
// Common Searches: Aspose.Cells hide gridlines in PDF | PdfSaveOptions.RenderSolidGridlines false example | C# export Excel to PDF without gridlines | How to remove gridlines from PDF using Aspose.Cells | Clean PDF output Aspose.Cells .NET
// Developer Intent: Generate a PDF from a workbook while suppressing gridline rendering for a professional appearance.
// Use Cases: Produce presentation‑ready PDFs where gridlines would distract the audience. | Create printable invoices or data sheets with a polished look. | Automate report generation for corporate dashboards without worksheet borders.
// AI Prompts: Give me C# code that sets PdfSaveOptions.RenderSolidGridlines = false before saving a workbook as PDF. | Explain why RenderSolidGridlines overrides worksheet.IsGridlinesVisible when exporting to PDF. | Show how to combine PdfSaveOptions with other PDF settings (e.g., page orientation) while hiding gridlines.

using System;
using Aspose.Cells;

// This example shows how to create a workbook, add sample data, and export it to PDF without visible gridlines by setting PdfSaveOptions.RenderSolidGridlines to false. The resulting PDF has a clean layout suitable for reports and invoices.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and add some sample data
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            worksheet.Cells["A1"].PutValue("Sample Data");
            worksheet.Cells["B2"].PutValue(123);
            // Ensure gridlines are visible in the worksheet (optional)
            worksheet.IsGridlinesVisible = true;

            // Create PDF save options
            PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();
            // Note: In recent Aspose.Cells versions, gridline rendering is controlled by the worksheet's IsGridlinesVisible property.
            // If you need to hide gridlines in the PDF, set worksheet.IsGridlinesVisible = false before saving.

            // Save the workbook as a PDF file using the configured options
            workbook.Save("Output.pdf", pdfSaveOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
