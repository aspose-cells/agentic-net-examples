// Title: Generate HTML from an Excel workbook with column letters using Aspose.Cells HtmlSaveOptions in C#
// AI Prompts: Write C# code that loads or creates a Workbook, enables HtmlSaveOptions.ExportColumnHeaders via reflection, and saves the workbook as HTML showing column letters (A, B, C…). | Show how to detect the ExportColumnHeaders property on HtmlSaveOptions with reflection and set it only when the property exists for backward‑compatible Aspose.Cells versions. | Provide a complete C# example that conditionally opens an existing .xlsx file, configures HTML export options, and outputs an HTML file that includes column headers.
// Common Searches: how to export Excel to HTML with column letters using Aspose.Cells C# | Aspose.Cells HtmlSaveOptions ExportColumnHeaders property example | C# reflection to enable ExportColumnHeaders in older Aspose.Cells versions | save workbook as HTML with column headers displayed Aspose.Cells | display column letters when converting Excel to HTML with Aspose.Cells
// Tags: Aspose.Cells HTML column header export | C# reflection for HtmlSaveOptions compatibility | Excel to HTML conversion with column letters | HtmlSaveOptions property detection Aspose.Cells | save workbook as HTML with column headers

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an existing workbook or creates a new one, uses reflection to enable the ExportColumnHeaders option in HtmlSaveOptions when available, and saves the workbook as an HTML file that displays column letters (A, B, C, …) alongside the data.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Path to an optional input workbook
                string inputPath = "input.xlsx";

                // Load an existing workbook if the file exists; otherwise create a new one
                Workbook workbook;
                if (File.Exists(inputPath))
                {
                    workbook = new Workbook(inputPath);
                }
                else
                {
                    workbook = new Workbook();
                }

                // Configure HTML save options
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);

                // The ExportColumnHeaders property may not be available in older versions.
                // If it exists, enable it to export column letters (A, B, C, ...).
                // This is done via reflection to maintain compatibility across versions.
                var prop = typeof(HtmlSaveOptions).GetProperty("ExportColumnHeaders");
                if (prop != null && prop.CanWrite)
                {
                    prop.SetValue(htmlOptions, true);
                }

                // Save the workbook as HTML
                string outputPath = "ExportedWithColumnHeaders.html";
                workbook.Save(outputPath, htmlOptions);

                Console.WriteLine($"Workbook successfully saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
