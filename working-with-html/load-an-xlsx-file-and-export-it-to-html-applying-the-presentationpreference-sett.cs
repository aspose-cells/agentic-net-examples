using System;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    class Program
    {
        static void Main()
        {
            // Path to the source XLSX file
            string sourcePath = "input.xlsx";

            // Load the workbook from the XLSX file
            Workbook workbook = new Workbook(sourcePath);

            // Create HTML save options
            HtmlSaveOptions saveOptions = new HtmlSaveOptions();

            // Enable presentation preference for a more beautiful layout
            saveOptions.PresentationPreference = true;

            // Export the workbook to HTML using the configured options
            workbook.Save("output.html", saveOptions);
        }
    }
}