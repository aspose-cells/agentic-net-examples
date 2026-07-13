using System;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    // Author: Example implementation for exporting comments to HTML
    class Program
    {
        static void Main()
        {
            // Load the existing XLS workbook
            Workbook workbook = new Workbook("input.xls");

            // Configure HTML save options to include all cell comments
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                IsExportComments = true // Export comments when saving to HTML
            };

            // Save the workbook as an HTML file with comments exported
            workbook.Save("output_with_comments.html", htmlOptions);
        }
    }
}