// Title: Set custom worksheet page margins and export to PDF using Aspose.Cells for .NET
// AI Prompts: Generate C# code that configures left, right, top, bottom, header, and footer margins on a worksheet via PageSetup and saves the workbook as a PDF with Aspose.Cells. | Demonstrate how to use PDF export options together with custom page margins to control the layout of a PDF generated from an Excel file in C# using Aspose.Cells.
// Common Searches: C# Aspose.Cells how to change page margins before PDF conversion | set side margins for an Excel worksheet when exporting to PDF with Aspose | adjust header/footer spacing in PDF output using Aspose.Cells .NET | PdfSaveOptions margin configuration example for Aspose.Cells | custom page layout settings for Excel to PDF conversion Aspose.Cells
// Tags: worksheet margin configuration Aspose.Cells | PDF layout control .NET | custom Excel to PDF margins | header footer spacing Aspose | export Excel with specific page layout

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsExample
{
    // The example creates a workbook, adds sample data, sets left, right, top, bottom, header, and footer margins via PageSetup, initializes PdfSaveOptions, and saves the file as CustomMargins.pdf.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (or load an existing one)
                Workbook workbook = new Workbook();

                // Access the first worksheet and add some sample data
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("Sample");
                sheet.Cells["A2"].PutValue("Data");

                // Set custom page margins (values are in inches)
                sheet.PageSetup.LeftMargin = 0.5;    // left margin
                sheet.PageSetup.RightMargin = 0.5;   // right margin
                sheet.PageSetup.TopMargin = 0.75;    // top margin
                sheet.PageSetup.BottomMargin = 0.75; // bottom margin
                sheet.PageSetup.HeaderMargin = 0.3;  // header margin
                sheet.PageSetup.FooterMargin = 0.3;  // footer margin

                // Create PdfSaveOptions to control PDF export settings (if additional options are needed)
                PdfSaveOptions pdfOptions = new PdfSaveOptions();

                // Save the workbook as a PDF using the configured options
                workbook.Save("CustomMargins.pdf", pdfOptions);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
