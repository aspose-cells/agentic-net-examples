// Title: Convert an Excel workbook to PDF with page scaling turned off using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# export Excel to PDF without fit-to-page scaling | How to keep original worksheet size when saving Excel as PDF with Aspose.Cells | Disable FitToPagesTall and FitToPagesWide for PDF conversion in Aspose.Cells | Set page zoom to 100% in Aspose.Cells PDF output | Preserve exact dimensions of Excel sheets in PDF using Aspose.Cells .NET
// Tags: Aspose.Cells PDF export without scaling | remove fit-to-page scaling Aspose.Cells | preserve worksheet dimensions PDF conversion | page zoom 100 percent Aspose.Cells | Excel to PDF exact size Aspose.Cells

using Aspose.Cells;
using System;

// The example loads an Excel workbook, disables FitToPagesTall and FitToPagesWide, sets the page zoom to 100% for each worksheet to prevent any scaling, and then saves the workbook as a PDF, preserving the original sheet dimensions exactly.
class Program
{
    static void Main()
    {
        // Load the source workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Ensure each worksheet uses no scaling (original dimensions)
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            PageSetup pageSetup = sheet.PageSetup;

            // Disable "Fit to page" scaling
            pageSetup.FitToPagesTall = 0;
            pageSetup.FitToPagesWide = 0;

            // Set zoom to 100% to keep the original size
            pageSetup.Zoom = 100;
        }

        // Convert and save the workbook as PDF
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
