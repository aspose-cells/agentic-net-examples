using System;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data
            sheet.Cells["A1"].PutValue("Name");
            sheet.Cells["B1"].PutValue("Age");
            sheet.Cells["A2"].PutValue("John");
            sheet.Cells["B2"].PutValue(30);
            sheet.Cells["A3"].PutValue("Alice");
            sheet.Cells["B3"].PutValue(25);

            // Configure HTML save options:
            // 1. DisableCss = true  -> no external CSS files, only inline styles are used.
            // 2. TableCssId = "custom-table" -> the generated <table> element will have
            //    an attribute TableCssId="custom-table" and all CSS class names will be prefixed
            //    with this value, allowing you to style the table via external CSS if desired.
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
            htmlOptions.DisableCss = true;
            htmlOptions.TableCssId = "custom-table";

            // Save the workbook as HTML using the configured options
            workbook.Save("ExportedTable.html", htmlOptions);
        }
    }
}