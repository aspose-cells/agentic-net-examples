using System;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source Excel file (XLSX)
            string inputPath = "input.xlsx";

            // Path for the resulting HTML file
            string outputPath = "output.html";

            // Load the workbook from the XLSX file
            Workbook workbook = new Workbook(inputPath);

            // Create HTML save options
            HtmlSaveOptions saveOptions = new HtmlSaveOptions();

            // Do not export hidden worksheets
            saveOptions.ExportHiddenWorksheet = false;

            // Save the workbook as HTML using the specified options
            workbook.Save(outputPath, saveOptions);

            Console.WriteLine($"Workbook saved to HTML without hidden worksheets: {outputPath}");
        }
    }
}