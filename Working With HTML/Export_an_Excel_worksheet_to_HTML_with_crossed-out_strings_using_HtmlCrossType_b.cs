using System;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    class Program
    {
        static void Main()
        {
            // Path to the source Excel file (XLSX)
            string inputPath = "input.xlsx";

            // Path where the resulting HTML file will be saved
            string outputPath = "output.html";

            // Load the workbook from the XLSX file
            Workbook workbook = new Workbook(inputPath);

            // Create HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

            // Set the cross-string display type.
            // HtmlCrossType.CrossHideRight will hide the right part of overlapping strings,
            // effectively showing a crossed-out appearance for cross‑cell strings.
            htmlOptions.HtmlCrossStringType = HtmlCrossType.CrossHideRight;

            // Save the workbook as HTML using the configured options
            workbook.Save(outputPath, htmlOptions);

            Console.WriteLine($"Workbook exported to HTML with cross‑string type at: {outputPath}");
        }
    }
}