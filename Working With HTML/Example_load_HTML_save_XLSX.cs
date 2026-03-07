using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsHtmlToXlsx
{
    class Program
    {
        static void Main()
        {
            // Path to the source HTML file
            string sourceHtml = "input.html";

            // Path for the resulting XLSX file
            string outputXlsx = "output.xlsx";

            // Create HTML load options (default constructor)
            HtmlLoadOptions loadOptions = new HtmlLoadOptions();

            // Load the HTML file into a Workbook using the load options
            Workbook workbook = new Workbook(sourceHtml, loadOptions);

            // Save the workbook as XLSX format
            workbook.Save(outputXlsx, SaveFormat.Xlsx);

            Console.WriteLine($"HTML file '{sourceHtml}' has been converted to '{outputXlsx}'.");
        }
    }
}