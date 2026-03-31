using System;
using Aspose.Cells;

namespace AsposeCellsHtmlConversion
{
    public class ConvertXlsxToHtml
    {
        public static void Run()
        {
            // Path to the source XLSX file
            string sourcePath = "input.xlsx";

            // Path where the resulting HTML file will be saved
            string outputPath = "output.html";

            // Load the workbook from the XLSX file
            Workbook workbook = new Workbook(sourcePath);

            // Create HTML save options
            HtmlSaveOptions saveOptions = new HtmlSaveOptions();

            // Enable exclusion of unused CSS styles to reduce file size
            saveOptions.ExcludeUnusedStyles = true;

            // Save the workbook as HTML using the specified options
            workbook.Save(outputPath, saveOptions);

            Console.WriteLine($"Workbook successfully converted to HTML at: {outputPath}");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ConvertXlsxToHtml.Run();
        }
    }
}