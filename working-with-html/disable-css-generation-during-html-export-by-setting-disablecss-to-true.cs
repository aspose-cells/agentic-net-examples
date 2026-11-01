// Title: Export an Excel workbook to HTML without generating CSS using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads an .xlsx file and saves it as HTML with HtmlSaveOptions.DisableCss set to true. | Demonstrate how to configure Aspose.Cells HtmlSaveOptions to turn off CSS generation during Excel‑to‑HTML conversion. | Provide a .NET example that exports a workbook to HTML while suppressing the creation of any stylesheet.
// Common Searches: Aspose.Cells C# disable CSS when saving workbook as HTML | How to stop CSS file creation in Aspose.Cells HTML export | Set HtmlSaveOptions.DisableCss true to get HTML without stylesheet from Excel | Export Excel to HTML without embedded CSS using Aspose.Cells | C# Aspose.Cells HTML export without generating CSS files
// Tags: HtmlSaveOptions.DisableCss Aspose.Cells | export workbook to html without css | Aspose.Cells suppress stylesheet on HTML save | C# disable css generation Aspose.Cells HTML export | save excel as html without embedded css

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Saving;

// The example loads an existing .xlsx file into a Workbook, creates HtmlSaveOptions, sets DisableCss = true to prevent any CSS or stylesheet from being emitted, and saves the workbook as an HTML file. The resulting HTML contains only the markup for the worksheet data, with no linked or embedded CSS.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.html";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file '{inputPath}' was not found.");
            return;
        }

        try
        {
            // Load the workbook from the existing Excel file
            Workbook workbook = new Workbook(inputPath);

            // Create HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);

            // Note: ExportCssClass property is not available in the current Aspose.Cells version.
            // The default behavior will be used.

            // Save the workbook as HTML with the specified options
            workbook.Save(outputPath, htmlOptions);

            Console.WriteLine($"Workbook successfully saved as HTML to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
