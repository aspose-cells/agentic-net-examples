using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.html";

            // Verify that the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook that already contains DataBar conditional formatting
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options to preserve DataBar appearance
            HtmlSaveOptions saveOptions = new HtmlSaveOptions();

            // Use a valid DataBarRenderMode value (BackgroundColor is supported in all versions)
            saveOptions.DataBarRenderMode = DataBarRenderMode.BackgroundColor;

            // Save the workbook as HTML using the configured options
            workbook.Save(outputPath, saveOptions);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}