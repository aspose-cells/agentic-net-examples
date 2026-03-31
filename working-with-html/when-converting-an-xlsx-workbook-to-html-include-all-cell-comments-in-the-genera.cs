using System;
using Aspose.Cells;

namespace AsposeCellsCommentExport
{
    class Program
    {
        static void Main()
        {
            // Load the source XLSX workbook
            Workbook workbook = new Workbook("input.xlsx");

            // Configure HTML save options to include cell comments
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                // Export all comments present in the workbook
                IsExportComments = true
            };

            // Save the workbook as an HTML file with comments embedded
            workbook.Save("output.html", htmlOptions);

            Console.WriteLine("Workbook converted to HTML with comments successfully.");
        }
    }
}