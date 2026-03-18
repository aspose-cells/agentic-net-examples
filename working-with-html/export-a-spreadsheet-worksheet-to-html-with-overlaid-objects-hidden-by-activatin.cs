using System;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    class Program
    {
        static void Main()
        {
            // Load the source XLSX workbook
            Workbook workbook = new Workbook("input.xlsx");

            // Configure HTML save options to hide overlapping text on the right side
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                HtmlCrossStringType = HtmlCrossType.CrossHideRight
            };

            // Save the workbook as HTML with the specified options
            workbook.Save("output.html", saveOptions);
        }
    }
}