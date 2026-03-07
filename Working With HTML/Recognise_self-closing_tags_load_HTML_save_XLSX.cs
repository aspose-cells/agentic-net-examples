using System;
using Aspose.Cells;

namespace AsposeCellsHtmlToXlsx
{
    class Program
    {
        static void Main()
        {
            // Path to the source HTML file that may contain self‑closing tags (e.g., <br/>, <img/>)
            string htmlPath = "input.html";

            // Create HTML load options.
            // SupportDivTag is enabled to correctly handle <div> layouts; other self‑closing tags are processed by default.
            HtmlLoadOptions loadOptions = new HtmlLoadOptions(LoadFormat.Html);
            loadOptions.SupportDivTag = true; // optional, helps with <div> handling

            // Load the HTML file into a workbook using the specified options.
            Workbook workbook = new Workbook(htmlPath, loadOptions);

            // Save the workbook as an XLSX file.
            string xlsxPath = "output.xlsx";
            workbook.Save(xlsxPath, SaveFormat.Xlsx);

            Console.WriteLine($"HTML file '{htmlPath}' has been converted to '{xlsxPath}'.");
        }
    }
}