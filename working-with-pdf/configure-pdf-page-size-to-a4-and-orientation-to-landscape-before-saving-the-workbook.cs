// Title: Aspose.Cells .NET: Set PDF page size to A4 and orientation to Landscape
// Description: Demonstrates how to create a workbook, configure the default paper size to A4 and the page orientation to landscape using WorkbookSettings and Worksheet PageSetup, then export the workbook as a PDF.
// Keywords: Aspose.Cells PDF page size | Aspose.Cells A4 | Aspose.Cells landscape orientation | .NET PDF export | WorkbookSettings PaperSize | PageSetup orientation
// Common Searches: set A4 page size for PDF export Aspose.Cells .NET | landscape orientation PDF Aspose.Cells workbook | how to change paper size before saving PDF with Aspose.Cells | configure page setup for all worksheets Aspose.Cells PDF
// Developer Intent: Configure the workbook’s PDF output to use A4 paper size and landscape orientation before saving.
// Use Cases: Produce printable reports in A4 landscape format directly from a spreadsheet. | Generate invoices or catalogs that require A4 landscape PDFs without manual adjustments. | Export charts and tables as ready‑to‑print A4 landscape PDFs for presentations.
// AI Prompts: Show code to set custom margins together with A4 landscape orientation in Aspose.Cells PDF export. | Explain how to apply different page orientations per worksheet when saving to PDF with Aspose.Cells. | Provide a sample that changes the paper size dynamically based on worksheet content in Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsPdfPageSetupDemo
{
    // Demonstrates how to create a workbook, configure the default paper size to A4 and the page orientation to landscape using WorkbookSettings and Worksheet PageSetup, then export the workbook as a PDF.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook (lifecycle create rule)
            Workbook workbook = new Workbook();

            // Set the default print paper size to A4 (WorkbookSettings.PaperSize property)
            workbook.Settings.PaperSize = PaperSizeType.PaperA4;

            // Set the print orientation for the whole workbook to Landscape
            // (WorkbookSettings.SetPageOrientationType method)
            workbook.Settings.SetPageOrientationType(PageOrientationType.Landscape);

            // Ensure each worksheet also has Landscape orientation (PageSetup.Orientation property)
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                sheet.PageSetup.Orientation = PageOrientationType.Landscape;
            }

            // Save the workbook as PDF (lifecycle save rule)
            workbook.Save("output.pdf", SaveFormat.Pdf);
        }
    }
}
