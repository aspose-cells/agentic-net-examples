using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the workbook from a file
        Workbook workbook = new Workbook("input.xlsx");

        // Enable page margins by setting them for each worksheet
        foreach (Worksheet ws in workbook.Worksheets)
        {
            // Margins are specified in inches
            ws.PageSetup.LeftMargin = 0.5;
            ws.PageSetup.RightMargin = 0.5;
            ws.PageSetup.TopMargin = 0.75;
            ws.PageSetup.BottomMargin = 0.75;
        }

        // Create HTML save options and embed images as Base64
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.ExportImagesAsBase64 = true;

        // Export the workbook to HTML with embedded images
        workbook.Save("output.html", htmlOptions);
    }
}