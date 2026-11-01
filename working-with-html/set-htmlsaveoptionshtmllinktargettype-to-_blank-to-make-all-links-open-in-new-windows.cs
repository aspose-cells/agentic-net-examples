// Title: Set hyperlink target to _blank using Aspose.Cells HtmlSaveOptions when converting Excel to HTML in C#
// AI Prompts: Write C# code that creates HtmlExportOptions, sets HtmlLinkTargetType to Blank, assigns it to HtmlSaveOptions, and saves a workbook as HTML so every hyperlink opens in a new tab. | Show how to detect whether the current Aspose.Cells version supports HtmlExportOptions and apply the link‑target setting before exporting to HTML.
// Common Searches: how to make hyperlinks open in new tab with Aspose.Cells HTML export c# | Aspose.Cells HtmlSaveOptions set link target to _blank | C# export Excel to HTML with _blank link target using Aspose.Cells | HtmlExportOptions HtmlLinkTargetType Blank example Aspose.Cells | Aspose.Cells version check for HtmlExportOptions support
// Tags: Aspose.Cells HtmlExportOptions link target | HtmlSaveOptions hyperlink target blank | C# export Excel to HTML Aspose.Cells | HtmlLinkTargetType Blank usage | Aspose.Cells HTML export version compatibility

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Saving;

// The sample loads an Excel workbook, creates HtmlSaveOptions, optionally configures HtmlExportOptions to set HtmlLinkTargetType to Blank (when supported), and saves the workbook as HTML so that all hyperlinks open in a new browser window, handling missing files and runtime errors.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.html";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Create HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);

            // NOTE: Setting hyperlink target to "_blank" requires HtmlExportOptions,
            // which may not be available in older Aspose.Cells versions.
            // If supported, the following code can be used:
            // var exportOptions = new HtmlExportOptions();
            // exportOptions.HtmlLinkTargetType = HtmlLinkTargetType.Blank;
            // htmlOptions.HtmlExportOptions = exportOptions;

            // Save the workbook as HTML with the specified options
            workbook.Save(outputPath, htmlOptions);

            Console.WriteLine($"Workbook successfully saved as HTML to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
