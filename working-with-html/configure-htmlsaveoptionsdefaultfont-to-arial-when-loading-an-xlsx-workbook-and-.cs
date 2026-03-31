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

            // Create HTML save options using the provided constructor
            HtmlSaveOptions saveOptions = new HtmlSaveOptions();

            // Set the default font name to Arial (property from the documentation)
            saveOptions.DefaultFontName = "Arial";

            // Export the workbook to HTML with the configured options
            workbook.Save("output.html", saveOptions);
        }
    }
}