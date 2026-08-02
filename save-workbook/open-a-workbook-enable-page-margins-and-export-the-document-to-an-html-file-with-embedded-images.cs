using System;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Enable and configure page margins for the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            // Margins are specified in points (1 point = 1/72 inch)
            sheet.PageSetup.LeftMargin = 36;   // 0.5 inch
            sheet.PageSetup.RightMargin = 36;  // 0.5 inch
            sheet.PageSetup.TopMargin = 36;    // 0.5 inch
            sheet.PageSetup.BottomMargin = 36; // 0.5 inch

            // Create HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

            // Export images as Base64 so they are embedded directly in the HTML file
            htmlOptions.ExportImagesAsBase64 = true;

            // Save the workbook as an HTML file with the specified options
            workbook.Save("output.html", htmlOptions);

            Console.WriteLine("Workbook exported to HTML with embedded images.");
        }
    }
}