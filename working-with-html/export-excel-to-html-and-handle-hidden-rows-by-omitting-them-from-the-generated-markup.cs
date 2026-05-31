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
            sheet.Cells["A1"].PutValue("Visible Row 1");
            sheet.Cells["A2"].PutValue("Hidden Row");
            sheet.Cells["A3"].PutValue("Visible Row 2");

            // Hide the second row (index 1)
            sheet.Cells.HideRow(1);

            // Configure HTML save options to remove hidden rows from the output
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
            htmlOptions.HiddenRowDisplayType = HtmlHiddenRowDisplayType.Remove;

            // Save the workbook as HTML; hidden rows will be omitted
            workbook.Save("ExportedWithoutHiddenRows.html", htmlOptions);

            Console.WriteLine("HTML file generated with hidden rows removed.");
        }
    }
}