// Title: How to disable CSS file generation when saving an Excel workbook as HTML with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# export Excel to HTML without creating a CSS file | set DisableCss true in HtmlSaveOptions to prevent external stylesheet Aspose.Cells | save workbook as HTML with inline styles only using Aspose.Cells .NET | C# Aspose.Cells HtmlSaveOptions disable CSS generation example | remove external CSS from Aspose.Cells HTML export in .NET
// Tags: HtmlSaveOptions.DisableCss | Aspose.Cells HTML export inline styling | C# save workbook as HTML without external CSS | prevent external stylesheet generation Aspose.Cells | minimal styling HTML save options

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Saving;

namespace AsposeCellsExample
{
    // The example loads an existing Excel file or creates a new workbook, configures HtmlSaveOptions with DisableCss set to true to suppress external CSS generation and embed minimal styling, and then saves the workbook as an HTML file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook (or load an existing one if the file exists)
                Workbook workbook;
                string inputPath = "input.xlsx";
                if (File.Exists(inputPath))
                {
                    workbook = new Workbook(inputPath);
                }
                else
                {
                    workbook = new Workbook(); // creates an empty workbook
                }

                // Configure HTML save options to disable CSS generation
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                {
                    // Embeds minimal styling directly into the HTML.
                    DisableCss = true
                };

                // Save the workbook as HTML with the specified options
                string outputPath = "output.html";
                workbook.Save(outputPath, htmlOptions);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
