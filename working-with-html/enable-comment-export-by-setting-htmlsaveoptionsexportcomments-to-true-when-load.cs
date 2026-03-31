using System;
using Aspose.Cells;

namespace ExportCommentsToHtml
{
    class Program
    {
        static void Main()
        {
            // Load the existing XLSX workbook
            Workbook workbook = new Workbook("input.xlsx");

            // Create HTML save options and enable comment export
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                IsExportComments = true
            };

            // Save the workbook as an HTML file with comments included
            workbook.Save("output.html", htmlOptions);
        }
    }
}