using System;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    // Author: Aspose.Cells .NET example – preserves worksheet name case in HTML headings
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet and set a custom name with mixed case
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "MyCustomSheet";

            // Add some sample data
            sheet.Cells["A1"].PutValue("Header1");
            sheet.Cells["B1"].PutValue("Header2");
            sheet.Cells["A2"].PutValue("Data1");
            sheet.Cells["B2"].PutValue("Data2");

            // Configure HTML save options to export row/column headings.
            // This ensures that the worksheet name appears in the generated HTML
            // using the exact capitalization defined above.
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                ExportRowColumnHeadings = true   // modern property (preferred over obsolete ExportHeadings)
                // ExportHeadings = true          // optional: kept for backward compatibility if needed
            };

            // Save the workbook as an HTML file; the heading tag will contain "MyCustomSheet"
            workbook.Save("MyCustomSheet.html", htmlOptions);
        }
    }
}