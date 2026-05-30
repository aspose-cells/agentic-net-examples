using System;
using Aspose.Cells;

namespace AsposeCellsExportHtmlWithGridlinesAndTableCssId
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Enable gridlines visibility in the worksheet
            worksheet.IsGridlinesVisible = true;

            // Add some sample data to demonstrate the output
            worksheet.Cells["A1"].PutValue("Header 1");
            worksheet.Cells["B1"].PutValue("Header 2");
            worksheet.Cells["A2"].PutValue("Value 1");
            worksheet.Cells["B2"].PutValue("Value 2");

            // Configure HTML save options
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                // Export the gridlines as they are visible in the worksheet
                ExportGridLines = worksheet.IsGridlinesVisible,

                // Apply a custom TableCssId for styling the generated HTML table
                TableCssId = "custom-table-style"
            };

            // Save the workbook as HTML using the configured options
            workbook.Save("ExportedWithGridlines.html", saveOptions);
        }
    }
}