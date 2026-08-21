// Title: Export a worksheet range to PDF with custom margins and landscape orientation using Aspose.Cells for .NET (C#)
// Description: Demonstrates how to set a print area (A1:C10), apply top, bottom, left, and right margins, switch to landscape orientation, and save the selected range as a PDF with PdfSaveOptions and optional CustomRenderSettings in C#.
// Keywords: Aspose.Cells | C# | PDF export | range to PDF | custom margins | page orientation | landscape | print area | PdfSaveOptions | CustomRenderSettings | Aspose.Cells for .NET
// Common Searches: Aspose.Cells export specific range to PDF | set PDF margins in Aspose.Cells C# | change page orientation to landscape when saving Excel as PDF Aspose | define print area before PDF conversion Aspose.Cells | use CustomRenderSettings with PdfSaveOptions Aspose
// Developer Intent: Save only the defined A1:C10 range as a PDF with specified margins and landscape layout.
// Use Cases: Create printable PDF reports that include only a portion of a worksheet. | Produce landscape‑oriented PDFs for wide tables or dashboards. | Apply precise margin settings to meet corporate printing standards. | Integrate custom rendering (e.g., scaling) via CustomRenderSettings before PDF generation. | Automate batch conversion of multiple sheet ranges to PDFs with consistent layout.
// AI Prompts: Write C# code using Aspose.Cells to export range A1:C10 to a PDF with 0.5" top/bottom and 0.75" left/right margins and landscape orientation. | Explain how to configure PdfSaveOptions and CustomRenderSettings to control margin and orientation when converting Excel to PDF with Aspose.Cells. | Show a loop that sets different print areas and margin values for several worksheets and saves each as a separate PDF file. | Provide troubleshooting steps if the PDF output does not reflect the custom margins or orientation. | Describe how to convert a range to PDF in Java or Python using Aspose.Cells equivalents.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsRangeToPdfDemo
{
    // Demonstrates how to set a print area (A1:C10), apply top, bottom, left, and right margins, switch to landscape orientation, and save the selected range as a PDF with PdfSaveOptions and optional CustomRenderSettings in C#.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data in the range A1:C10
            for (int row = 0; row < 10; row++)
            {
                for (int col = 0; col < 3; col++)
                {
                    sheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // Define the range to be rendered as PDF
            string printArea = "A1:C10";
            sheet.PageSetup.PrintArea = printArea;

            // Apply custom margins (in inches)
            sheet.PageSetup.TopMarginInch = 0.5f;
            sheet.PageSetup.BottomMarginInch = 0.5f;
            sheet.PageSetup.LeftMarginInch = 0.75f;
            sheet.PageSetup.RightMarginInch = 0.75f;

            // Set page orientation (Landscape or Portrait)
            sheet.PageSetup.Orientation = PageOrientationType.Landscape;

            // Create PDF save options (inherits from PaginatedSaveOptions)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Optional: assign a CustomRenderSettings instance if further custom rendering is needed
            pdfOptions.CustomRenderSettings = new CustomRenderSettings();

            // Save the workbook as PDF; only the defined print area will be rendered
            workbook.Save("RangeRendered.pdf", pdfOptions);

            Console.WriteLine("Range rendered to PDF with custom margins and orientation.");
        }
    }
}
