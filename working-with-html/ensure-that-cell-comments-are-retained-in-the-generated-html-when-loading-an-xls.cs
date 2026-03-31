using System;
using Aspose.Cells;

namespace AsposeCellsCommentExport
{
    class Program
    {
        static void Main()
        {
            // Load the existing XLSX workbook
            Workbook workbook = new Workbook("input.xlsx");

            // Configure HTML save options to export cell comments
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                // Enable exporting of comments; default is false
                IsExportComments = true
            };

            // Save the workbook as HTML with comments retained
            workbook.Save("output.html", htmlOptions);

            Console.WriteLine("HTML file generated with cell comments exported.");
        }
    }
}