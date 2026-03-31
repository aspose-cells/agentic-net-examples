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

            // Create HTML save options and enable simulation of unsupported borders
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                // When true, Aspose.Cells will export a similar border style for borders
                // that are not supported by browsers, preserving the visual appearance.
                ExportSimilarBorderStyle = true
            };

            // Save the workbook as HTML using the configured options
            workbook.Save("output.html", htmlOptions);
        }
    }
}