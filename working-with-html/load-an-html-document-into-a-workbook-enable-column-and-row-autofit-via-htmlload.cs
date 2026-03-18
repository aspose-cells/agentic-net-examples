using System;
using Aspose.Cells;

namespace AsposeCellsHtmlImport
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source HTML file
            string htmlPath = "input.html";

            // Create HtmlLoadOptions and enable auto‑fit for columns and rows
            HtmlLoadOptions loadOptions = new HtmlLoadOptions(LoadFormat.Html);
            loadOptions.AutoFitColsAndRows = true;

            // Load the HTML file into a workbook using the specified options
            Workbook workbook = new Workbook(htmlPath, loadOptions);

            // Path for the resulting XLSX file
            string xlsxPath = "output.xlsx";

            // Save the workbook as an XLSX file
            workbook.Save(xlsxPath, SaveFormat.Xlsx);

            Console.WriteLine($"HTML file '{htmlPath}' has been loaded with auto‑fit applied and saved as '{xlsxPath}'.");
        }
    }
}