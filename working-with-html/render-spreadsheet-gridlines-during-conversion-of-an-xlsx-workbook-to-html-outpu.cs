using System;
using Aspose.Cells;

class RenderGridlinesToHtml
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Ensure gridlines are visible in the worksheet
        worksheet.IsGridlinesVisible = true;

        // Add some sample data to demonstrate the gridlines
        worksheet.Cells["A1"].PutValue("Header");
        worksheet.Cells["A2"].PutValue("Item 1");
        worksheet.Cells["B2"].PutValue(100);
        worksheet.Cells["A3"].PutValue("Item 2");
        worksheet.Cells["B3"].PutValue(200);

        // Configure HTML save options to export gridlines
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions
        {
            // Export gridlines based on the worksheet setting
            ExportGridLines = worksheet.IsGridlinesVisible,
            // Export only the active worksheet for simplicity
            ExportActiveWorksheetOnly = true
        };

        // Save the workbook as an HTML file with gridlines rendered
        workbook.Save("output.html", htmlOptions);
    }
}