using System;
using Aspose.Cells;

class ExportPrintAreaToHtml
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate the worksheet with sample data
        for (int row = 0; row < 20; row++)
        {
            for (int col = 0; col < 10; col++)
            {
                worksheet.Cells[row, col].PutValue($"Cell {row + 1},{col + 1}");
            }
        }

        // Define the print area that should be exported
        worksheet.PageSetup.PrintArea = "B2:F10";

        // Set HTML save options to export only the defined print area
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.ExportPrintAreaOnly = true;   // Export only the print area
        htmlOptions.ExportGridLines = true;       // Optional: include grid lines in the HTML

        // Save the workbook as an HTML file; only the print area will appear
        workbook.Save("PrintAreaOnly.html", htmlOptions);
    }
}