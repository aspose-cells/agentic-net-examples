using System;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    // Author: Aspose.Cells .NET example – export worksheets with titles as headings
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // ----- Worksheet 1 -----
            Worksheet sheet1 = workbook.Worksheets[0]; // default first sheet
            sheet1.Name = "Sales Summary";
            sheet1.Cells["A1"].PutValue("Product");
            sheet1.Cells["B1"].PutValue("Quantity");
            sheet1.Cells["A2"].PutValue("Apples");
            sheet1.Cells["B2"].PutValue(150);
            sheet1.Cells["A3"].PutValue("Oranges");
            sheet1.Cells["B3"].PutValue(200);

            // ----- Worksheet 2 -----
            int sheet2Index = workbook.Worksheets.Add();
            Worksheet sheet2 = workbook.Worksheets[sheet2Index];
            sheet2.Name = "Inventory";
            sheet2.Cells["A1"].PutValue("Item");
            sheet2.Cells["B1"].PutValue("Stock");
            sheet2.Cells["A2"].PutValue("Pens");
            sheet2.Cells["B2"].PutValue(500);
            sheet2.Cells["A3"].PutValue("Notebooks");
            sheet2.Cells["B3"].PutValue(300);

            // Configure HTML save options to export worksheet titles as top‑level headings
            HtmlSaveOptions options = new HtmlSaveOptions
            {
                // ExportRowColumnHeadings also adds the worksheet name as a heading in the HTML output
                ExportRowColumnHeadings = true,

                // Optional: export all worksheets into a single HTML file
                ExportActiveWorksheetOnly = false
            };

            // Save the workbook as a single HTML file; each worksheet name appears as a heading
            workbook.Save("WorksheetsWithHeadings.html", options);
        }
    }
}