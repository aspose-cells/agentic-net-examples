using System;
using Aspose.Cells;

namespace PreserveWorksheetNameCapitalization
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet and give it a mixed‑case name
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "SalesReport2023";   // original capitalization to be preserved

            // Add some sample data
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Quantity");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(150);
            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["B3"].PutValue(200);

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                // Use the worksheet name as the HTML page title – this keeps the exact case
                PageTitle = sheet.Name,

                // Export the sheet name as a heading inside the HTML file
                ExportPageHeaders = true,

                // Save all sheets into a single HTML file (optional, but keeps headings together)
                SaveAsSingleFile = true,
                ShowAllSheets = true
            };

            // Save the workbook as HTML; the heading tags will contain "SalesReport2023"
            workbook.Save("SalesReport2023.html", htmlOptions);
        }
    }
}