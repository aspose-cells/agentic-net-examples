using System;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    // Author: Aspose.Cells .NET example
    class ExportWorksheetHeadings
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Rename the default worksheet
            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.Name = "SalesData";

            // Add a second worksheet and set its name
            int sheet2Index = workbook.Worksheets.Add();
            Worksheet sheet2 = workbook.Worksheets[sheet2Index];
            sheet2.Name = "Summary";

            // Populate some sample data in the first sheet
            sheet1.Cells["A1"].PutValue("Product");
            sheet1.Cells["B1"].PutValue("Quantity");
            sheet1.Cells["A2"].PutValue("Apple");
            sheet1.Cells["B2"].PutValue(120);
            sheet1.Cells["A3"].PutValue("Banana");
            sheet1.Cells["B3"].PutValue(85);

            // Populate some sample data in the second sheet
            sheet2.Cells["A1"].PutValue("Report");
            sheet2.Cells["A2"].PutValue("Generated on:");
            sheet2.Cells["B2"].PutValue(DateTime.Now.ToString("yyyy-MM-dd"));

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                // Export row and column headings (e.g., A, B, 1, 2)
                ExportRowColumnHeadings = true,

                // Export worksheet properties so that each worksheet name appears as an HTML heading
                ExportWorksheetProperties = true,

                // Optional: export only the active worksheet (set to false to include all)
                ExportActiveWorksheetOnly = false
            };

            // Save the workbook as an HTML file with the configured options
            workbook.Save("WorkbookWithHeadings.html", htmlOptions);
        }
    }
}