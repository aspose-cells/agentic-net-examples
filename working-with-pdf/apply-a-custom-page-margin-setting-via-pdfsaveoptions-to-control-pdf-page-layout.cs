using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Sample data (optional)
        sheet.Cells["A1"].PutValue("Custom Page Margin Demo");
        sheet.Cells["A2"].PutValue("Margins are set via PageSetup properties.");

        // Set custom page margins (values are in inches)
        sheet.PageSetup.LeftMargin   = 0.5; // left margin
        sheet.PageSetup.RightMargin  = 0.5; // right margin
        sheet.PageSetup.TopMargin    = 1.0; // top margin
        sheet.PageSetup.BottomMargin = 1.0; // bottom margin

        // Create PdfSaveOptions to control PDF saving
        PdfSaveOptions saveOptions = new PdfSaveOptions();

        // Save the workbook as PDF using the custom margin settings
        workbook.Save("CustomMargins.pdf", saveOptions);
    }
}

// Author: Aspose.Cells .NET example demonstrating custom PDF page margins.