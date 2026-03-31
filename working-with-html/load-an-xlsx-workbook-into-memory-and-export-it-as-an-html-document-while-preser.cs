using System;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source Excel file (XLSX)
            string sourcePath = "input.xlsx";

            // Load the workbook into memory
            Workbook workbook = new Workbook(sourcePath);

            // Create HTML save options – default settings preserve formatting
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

            // Optional: ensure formulas are calculated before export
            htmlOptions.CalculateFormula = true;

            // Export the entire workbook as an HTML document
            string outputPath = "output.html";
            workbook.Save(outputPath, htmlOptions);

            Console.WriteLine($"Workbook successfully exported to HTML: {outputPath}");
        }
    }
}