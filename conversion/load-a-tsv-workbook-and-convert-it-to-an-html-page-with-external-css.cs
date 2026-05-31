using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsTsvToHtml
{
    class Program
    {
        static void Main()
        {
            // Path to the source TSV file
            string tsvPath = Path.Combine(Environment.CurrentDirectory, "input.tsv");

            // Load the TSV workbook (Aspose.Cells detects the format automatically)
            Workbook workbook = new Workbook(tsvPath);

            // Configure HTML save options to export CSS to a separate file
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                ExportWorksheetCSSSeparately = true   // CSS will be written to a separate .css file
            };

            // Define output HTML file path
            string htmlPath = Path.Combine(Environment.CurrentDirectory, "output.html");

            // Save the workbook as HTML with the specified options
            workbook.Save(htmlPath, htmlOptions);

            Console.WriteLine("TSV file has been converted to HTML.");
            Console.WriteLine($"HTML file: {htmlPath}");
            Console.WriteLine("A separate CSS file has been generated alongside the HTML.");
        }
    }
}