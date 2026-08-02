// Title: Export Cell Comments to HTML with Aspose.Cells for .NET (C#)
// Description: Loads an XLS workbook, sets HtmlSaveOptions.IsExportComments to true, and saves the file as HTML so every cell comment appears in the generated web page.
// Keywords: Aspose.Cells export comments | HTML output with cell notes | IsExportComments C# | XLS to HTML conversion | Aspose.Cells .NET example
// Common Searches: Aspose.Cells include comments in HTML export | C# HtmlSaveOptions IsExportComments usage | convert Excel file to HTML with comments | save workbook as HTML preserving notes | Aspose.Cells export cell annotations
// Developer Intent: Create an HTML version of an XLS workbook that retains all cell comments.
// Use Cases: Display annotated Excel reports directly in a web portal. | Provide a preview of uploaded spreadsheets with comments for end‑users. | Automate nightly conversion of Excel dashboards into HTML emails while keeping remarks.
// AI Prompts: Generate C# code using Aspose.Cells to load an XLS file, enable IsExportComments, and save it as HTML. | Describe how the IsExportComments flag influences the HTML markup and where comments are rendered. | Recommend best‑practice error handling for exporting Excel comments to HTML with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Loads an XLS workbook, sets HtmlSaveOptions.IsExportComments to true, and saves the file as HTML so every cell comment appears in the generated web page.
    public class ExportCommentsToHtml
    {
        public static void Run()
        {
            const string inputPath = "input.xls";
            const string outputPath = "output.html";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                throw new FileNotFoundException($"The input file '{inputPath}' was not found.");
            }

            try
            {
                // Load the existing XLS workbook
                Workbook workbook = new Workbook(inputPath);

                // Set HTML save options to include all cell comments
                HtmlSaveOptions saveOptions = new HtmlSaveOptions
                {
                    IsExportComments = true
                };

                // Save the workbook as HTML with comments exported
                workbook.Save(outputPath, saveOptions);

                Console.WriteLine($"Workbook successfully exported to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Log any runtime errors
                Console.WriteLine($"An error occurred while exporting comments: {ex.Message}");
                throw;
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            try
            {
                ExportCommentsToHtml.Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }
    }
}
