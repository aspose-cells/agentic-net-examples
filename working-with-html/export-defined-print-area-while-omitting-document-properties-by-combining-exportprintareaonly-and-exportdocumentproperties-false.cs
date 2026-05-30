using System;
using Aspose.Cells;

class ExportPrintAreaWithoutDocumentProperties
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

        // Set up HTML save options:
        // - ExportPrintAreaOnly = true  => only the defined print area will be saved
        // - ExportDocumentProperties = false => document properties are omitted from the output
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions
        {
            ExportPrintAreaOnly = true,
            ExportDocumentProperties = false
        };

        // Save the workbook as HTML using the configured options
        workbook.Save("PrintAreaWithoutDocProps.html", htmlOptions);
    }
}