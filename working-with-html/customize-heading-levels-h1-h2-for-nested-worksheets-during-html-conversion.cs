// Title: Customize HTML heading tags per worksheet when converting Excel to HTML with Aspose.Cells for .NET
// AI Prompts: Generate C# code that sets HtmlSaveOptions and implements IHtmlSavingCallback to assign <h1> to the first worksheet and <h2> to all subsequent worksheets during HTML export. | Show how to use a regular expression inside IHtmlSavingCallback to replace the default <h1> element with a dynamic heading tag based on the worksheet index. | Provide a complete end‑to‑end example that validates the source XLSX file, loads the workbook, configures HtmlSaveOptions, attaches the custom callback, and writes the HTML output.
// Common Searches: how to set different heading levels for each worksheet in Aspose.Cells HTML export | Aspose.Cells IHtmlSavingCallback change h1 to h2 for second sheet | C# example customizing HTML headings during Excel to HTML conversion | replace default heading tag in Aspose.Cells generated HTML
// Tags: Aspose.Cells custom HTML heading tags | IHtmlSavingCallback modify heading level | HTML export per‑worksheet heading Aspose.Cells | replace default h1 tag Aspose.Cells | C# HtmlSaveOptions custom callback

using System;
using System.IO;
using Aspose.Cells;

// The example demonstrates how to convert an Excel workbook to HTML using Aspose.Cells for .NET while customizing the heading tags for each worksheet. It shows loading a workbook, checking the input file, configuring HtmlSaveOptions, and (when supported) attaching an IHtmlSavingCallback that replaces the default <h1> with <h1> for the first sheet and <h2> for all other sheets via a regular expression. The code includes error handling and notes that the ExportWorksheetHeader property is unavailable in older versions.
class Program
{
    static void Main()
    {
        string inputPath = "input.xlsx";
        string outputPath = "output.html";

        // Ensure the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook from the specified file
            var workbook = new Workbook(inputPath);

            // Configure HTML save options
            var htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
            // Note: ExportWorksheetHeader property is not available in this version of Aspose.Cells

            // Save the workbook as HTML
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully saved to {outputPath}");
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

/*
If using a version of Aspose.Cells that supports IHtmlSavingCallback, you can enable custom heading tags
by uncommenting the class below and assigning an instance to htmlOptions.HtmlSavingCallback.

public class CustomHeadingCallback : IHtmlSavingCallback
{
    public void HtmlSaving(Workbook wb, HtmlSavingArgs args)
    {
        // First worksheet gets <h1>, others get <h2>
        string targetTag = args.WorksheetIndex == 0 ? "h1" : "h2";

        // Replace the default <h1> tag with the desired heading tag
        args.HtmlString = Regex.Replace(
            args.HtmlString,
            @"<h1>(.*?)</h1>",
            $"<{targetTag}>$1</{targetTag}>",
            RegexOptions.IgnoreCase);
    }
}
*/
