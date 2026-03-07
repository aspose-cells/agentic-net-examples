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

            // Create HTML save options and set the cross-string handling to hide the right overlapping text
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                HtmlCrossStringType = HtmlCrossType.CrossHideRight
            };

            // Save the workbook as an HTML file using the configured options
            workbook.Save("output.html", htmlOptions);

            Console.WriteLine("Workbook has been saved to HTML with HtmlCrossStringType = CrossHideRight.");
        }
    }
}