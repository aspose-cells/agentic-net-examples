using System;
using Aspose.Cells;

namespace AsposeCellsHtmlImport
{
    class Program
    {
        static void Main()
        {
            // Path to the source HTML file
            string htmlPath = "input.html";

            // Create HtmlLoadOptions (self‑closing tags are handled automatically)
            HtmlLoadOptions loadOptions = new HtmlLoadOptions();

            // Load the HTML file into a Workbook using the specified options
            Workbook workbook = new Workbook(htmlPath, loadOptions);

            // Save the resulting workbook as an XLSX file
            workbook.Save("output.xlsx");
        }
    }
}