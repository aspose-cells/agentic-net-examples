using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class ExcludeUnusedStylesDemo
    {
        public static void Run()
        {
            // Path to the source XLSX file
            string inputPath = "input.xlsx";

            // Load the workbook from the XLSX file
            Workbook workbook = new Workbook(inputPath);

            // Remove any styles that are not used in the workbook.
            workbook.RemoveUnusedStyles();

            // Create HTML save options
            HtmlSaveOptions saveOptions = new HtmlSaveOptions();

            // Ensure that unused CSS styles are excluded from the generated HTML.
            saveOptions.ExcludeUnusedStyles = true;

            // Path for the output HTML file
            string outputPath = "output.html";

            // Save the workbook as HTML using the configured options
            workbook.Save(outputPath, saveOptions);

            Console.WriteLine($"Workbook saved to HTML with unused styles excluded: {outputPath}");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ExcludeUnusedStylesDemo.Run();
        }
    }
}