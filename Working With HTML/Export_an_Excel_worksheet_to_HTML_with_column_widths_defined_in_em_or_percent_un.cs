using System;
using Aspose.Cells;

namespace ExportExcelToHtmlWithColumnWidthUnit
{
    class Program
    {
        static void Main()
        {
            // Load the existing XLSX workbook from disk
            Workbook workbook = new Workbook("input.xlsx");

            // Create HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

            // If your Aspose.Cells version supports it, you can set the column width unit like this:
            // htmlOptions.ColumnWidthUnit = HtmlColumnWidthUnit.Em;

            // Save the workbook as an HTML file using the configured options
            workbook.Save("output.html", htmlOptions);
        }
    }
}