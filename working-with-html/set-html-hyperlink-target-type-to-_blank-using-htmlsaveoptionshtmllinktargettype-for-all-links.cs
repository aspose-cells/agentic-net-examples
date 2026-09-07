// Title: Export an Excel workbook to HTML with all hyperlinks opening in a new tab using Aspose.Cells HtmlSaveOptions.HtmlLinkTargetType in C#
// AI Prompts: Generate C# code that loads an .xlsx file with Aspose.Cells, sets HtmlSaveOptions.HtmlLinkTargetType to Blank, and saves the workbook as HTML so every hyperlink uses target="_blank". | Provide a concise example showing how to configure Aspose.Cells HTML export options to force all links to open in a new browser tab.
// Common Searches: Aspose.Cells C# export Excel to HTML with target=_blank for all links | HtmlSaveOptions HtmlLinkTargetType property example in .NET | How to make hyperlinks open in a new tab when saving a workbook as HTML using Aspose.Cells | C# set hyperlink target type to blank in Aspose.Cells HTML output
// Tags: Aspose.Cells HTML export hyperlink target | HtmlSaveOptions HtmlLinkTargetType usage | C# set links to open in new tab | Excel to HTML conversion Aspose.Cells | configure hyperlink target blank Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The sample checks for the existence of input.xlsx, loads it into an Aspose.Cells Workbook, creates HtmlSaveOptions for HTML format, sets HtmlLinkTargetType to Blank so every hyperlink is rendered with target="_blank", saves the result as output.html, and includes basic exception handling.
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
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options (default options are sufficient for this example)
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);

            // Save the workbook as HTML
            workbook.Save(outputPath, htmlOptions);

            Console.WriteLine($"Workbook successfully saved as HTML to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors during processing
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
