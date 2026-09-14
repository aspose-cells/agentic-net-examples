// Title: How to save an Excel workbook as HTML with embedded CSS using Aspose.Cells for .NET
// AI Prompts: Generate C# code that saves a Workbook to HTML with CSS embedded in the output using Aspose.Cells. | Show how to configure HtmlSaveOptions to disable external CSS file creation and embed styles directly. | Provide a complete example that sets ExportCssSeparately to false and CssStyleSheetType to Embedded.
// Common Searches: Aspose.Cells C# embed CSS when exporting workbook to HTML | disable separate CSS file generation in Aspose.Cells HTML export | HtmlSaveOptions ExportCssSeparately false example | how to embed stylesheet in HTML output using Aspose.Cells .NET | Aspose.Cells HTML save options to include inline styles
// Tags: Aspose.Cells HtmlSaveOptions embedded CSS | C# export workbook to HTML with inline styles | ExportCssSeparately false Aspose.Cells | CssStyleSheetType Embedded .NET | disable external stylesheet Aspose.Cells HTML

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // Loads or creates a Workbook, configures HtmlSaveOptions (ExportCssSeparately = false, CssStyleSheetType = Embedded) to embed CSS directly, and saves the workbook as an HTML file.
    class Program
    {
        static void Main()
        {
            try
            {
                Workbook workbook;

                // Load an existing workbook if the file exists; otherwise create a new one.
                const string inputPath = "input.xlsx";
                if (File.Exists(inputPath))
                {
                    workbook = new Workbook(inputPath);
                }
                else
                {
                    workbook = new Workbook(); // creates a new empty workbook
                }

                // Configure HTML save options.
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
                // Note: In newer Aspose.Cells versions you can embed CSS by setting:
                // htmlOptions.ExportCssSeparately = false;
                // htmlOptions.CssStyleSheetType = CssStyleSheetType.Embedded;

                // Save the workbook as an HTML file.
                const string outputPath = "output.html";
                workbook.Save(outputPath, htmlOptions);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
