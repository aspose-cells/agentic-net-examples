using System;
using Aspose.Cells;

namespace AsposeCellsExportHiddenWorksheetDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source XLSX workbook
            string inputPath = "input.xlsx";

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Create HTML save options
            HtmlSaveOptions saveOptions = new HtmlSaveOptions();

            // Disable exporting of hidden worksheets
            saveOptions.ExportHiddenWorksheet = false;

            // Path for the resulting HTML file
            string outputPath = "output_without_hidden.html";

            // Save the workbook as HTML using the configured options
            workbook.Save(outputPath, saveOptions);

            Console.WriteLine($"Workbook exported to HTML without hidden worksheets: {outputPath}");
        }
    }
}