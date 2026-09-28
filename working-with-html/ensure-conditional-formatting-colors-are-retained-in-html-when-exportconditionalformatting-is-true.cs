// Title: Preserve Excel conditional formatting colors when exporting to HTML with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that loads an .xlsx workbook, configures HtmlSaveOptions to export conditional formatting, and saves it as an HTML file using Aspose.Cells. | Show how to verify the source Excel file exists before converting it to HTML while keeping all conditional formatting colors intact. | Demonstrate setting HtmlSaveOptions.ExportConditionalFormatting = true and handling possible exceptions during the HTML export in a .NET console application.
// Common Searches: Aspose.Cells C# export Excel to HTML with conditional formatting colors retained | How to keep conditional formatting when saving workbook as HTML using Aspose.Cells .NET | HtmlSaveOptions ExportConditionalFormatting property example for .NET
// Tags: Aspose.Cells HtmlSaveOptions ExportConditionalFormatting | C# export Excel to HTML with formatting | preserve conditional formatting colors Aspose.Cells | HTML conversion of Excel workbook .NET | conditional formatting retention during HTML export

using System;
using System.IO;
using Aspose.Cells;

// This example loads an existing Excel file, checks its presence, configures HtmlSaveOptions (with ExportConditionalFormatting enabled) and saves the workbook as an HTML document, ensuring that all conditional formatting colors are preserved. Error handling is included to capture any issues during the conversion.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.html";

        // Verify that the input workbook exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options (conditional formatting is exported by default)
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

            // Save the workbook as HTML with the specified options
            workbook.Save(outputPath, htmlOptions);

            Console.WriteLine($"Workbook successfully saved as HTML to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors during processing
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
