// Title: C# – Export Excel to PDF with Solid Gridlines using Aspose.Cells PdfSaveOptions
// Description: Shows how to enable worksheet gridlines, set PdfSaveOptions.GridlineType to Hair (solid), and save a workbook as a PDF with solid gridlines using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | PdfSaveOptions | GridlineType.Hair | solid gridlines PDF | C# Excel to PDF | RenderSolidGridlines | gridlines visibility | Aspose.Cells PDF export | Aspose.Cells .NET | Excel PDF gridlines
// Common Searches: Aspose.Cells export Excel to PDF with solid gridlines | PdfSaveOptions GridlineType Hair example C# | How to render gridlines when saving workbook as PDF using Aspose.Cells | Enable solid gridlines in PDF output with Aspose.Cells .NET | C# code to save workbook to PDF with visible gridlines
// Developer Intent: Create a PDF from an Excel workbook where the gridlines appear as solid lines.
// Use Cases: Printable reports that require clear cell separation | PDF invoices preserving the original spreadsheet layout | Technical documentation exported to PDF while keeping the grid structure
// AI Prompts: Provide a C# snippet that sets PdfSaveOptions.GridlineType to Hair for solid gridlines in Aspose.Cells. | How can I customize the gridline color when exporting an Excel sheet to PDF with Aspose.Cells? | Explain the effect of different GridlineType values (Hair, Dotted, Dashed) on PDF exports in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsExamples
{
    // Shows how to enable worksheet gridlines, set PdfSaveOptions.GridlineType to Hair (solid), and save a workbook as a PDF with solid gridlines using Aspose.Cells for .NET.
    public class RenderSolidGridlinesPdfDemo
    {
        // Entry point required for console application
        public static void Main(string[] args)
        {
            try
            {
                Run();
                Console.WriteLine("PDF generated successfully: SolidGridlinesDemo.pdf");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add some sample data so that gridlines are visible
            sheet.Cells["A1"].PutValue("Solid Gridlines Demo");
            sheet.Cells["B2"].PutValue(123);
            sheet.Cells["C3"].PutValue(DateTime.Now);

            // Enable gridlines visibility in the worksheet
            sheet.IsGridlinesVisible = true;

            // Create PDF save options and set gridline rendering to solid (Hair)
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                GridlineType = GridlineType.Hair
                // Optional: customize gridline color
                // GridlineColor = System.Drawing.Color.Black
            };

            // Save the workbook as PDF with the specified options
            workbook.Save("SolidGridlinesDemo.pdf", pdfOptions);
        }
    }
}
