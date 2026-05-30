using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class ExportWorksheetCssSeparatelyDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and add sample data
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("Hello");
                sheet.Cells["B2"].PutValue(123);

                // Initialize HtmlSaveOptions and embed CSS within the HTML
                HtmlSaveOptions saveOptions = new HtmlSaveOptions
                {
                    ExportWorksheetCSSSeparately = false // embed CSS
                };

                // Define output HTML file path
                string outputPath = Path.Combine(Path.GetTempPath(), "ExportWorksheetCssSeparately.html");

                // Save the workbook as HTML using the specified options
                workbook.Save(outputPath, saveOptions);

                // Verify that the HTML file was created
                if (!File.Exists(outputPath))
                {
                    Console.WriteLine($"Failed to create HTML file at: {outputPath}");
                    return;
                }

                // Read the generated HTML content
                string htmlContent = File.ReadAllText(outputPath);

                // Verify that CSS is embedded within the HTML (look for a <style> tag)
                bool isCssEmbedded = htmlContent.IndexOf("<style", StringComparison.OrdinalIgnoreCase) >= 0;

                // Output verification result
                Console.WriteLine($"HTML file saved to: {outputPath}");
                Console.WriteLine($"CSS embedded within HTML: {isCssEmbedded}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Runtime error: {ex.Message}");
            }
        }

        // Entry point for the application
        public static void Main(string[] args)
        {
            Run();
        }
    }
}