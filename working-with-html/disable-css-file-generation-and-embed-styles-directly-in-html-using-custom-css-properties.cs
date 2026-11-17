// Title: Generate HTML from an Excel workbook with inline CSS using Aspose.Cells for .NET (disable external stylesheet)
// Common Searches: how to stop Aspose.Cells from creating a separate CSS file when saving to HTML in C# | Aspose.Cells export Excel to HTML with styles embedded in the markup | C# convert .xlsx to HTML with inline CSS using Aspose.Cells HtmlSaveOptions | disable external stylesheet generation in Aspose.Cells HTML export
// Tags: Aspose.Cells HtmlSaveOptions embed inline styles | C# export Excel to HTML without external CSS file | disable CSS file generation Aspose.Cells | inline CSS properties in Aspose.Cells HTML output | Excel to HTML conversion with embedded styling C#

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, configures HtmlSaveOptions to embed style information directly in the generated HTML, and saves the file without producing a separate CSS stylesheet.
class Program
{
    static void Main()
    {
        // Define input and output file paths
        string inputPath = @"C:\Path\To\YourWorkbook.xlsx";
        string outputPath = @"C:\Path\To\Output.html";

        // Verify that the input workbook exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options (default behavior embeds styles)
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);

            // Save the workbook as HTML
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully saved to HTML: {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any errors that occur during loading or saving
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
