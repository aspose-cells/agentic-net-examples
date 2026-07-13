using System;
using System.IO;
using Aspose.Cells;

// Author: Aspose.Cells .NET example – compares HTML output before and after disabling downlevel‑revealed comments
class Program
{
    static void Main()
    {
        // Create a sample workbook with a comment
        Workbook workbook = new Workbook();                     // create workbook
        Worksheet worksheet = workbook.Worksheets[0];
        int commentIdx = worksheet.Comments.Add("A1");
        Comment comment = worksheet.Comments[commentIdx];
        comment.Note = "This is a test comment";

        // Save with default HtmlSaveOptions (DisableDownlevelRevealedComments = false)
        HtmlSaveOptions defaultOptions = new HtmlSaveOptions(); // default options
        string defaultHtmlPath = "output_default.html";
        workbook.Save(defaultHtmlPath, defaultOptions);         // save workbook

        // Read the generated HTML
        string defaultHtml = File.ReadAllText(defaultHtmlPath);

        // Save with DisableDownlevelRevealedComments set to true
        HtmlSaveOptions disabledOptions = new HtmlSaveOptions();
        disabledOptions.DisableDownlevelRevealedComments = true; // disable downlevel‑revealed comments
        string disabledHtmlPath = "output_disabled.html";
        workbook.Save(disabledHtmlPath, disabledOptions);       // save workbook

        // Read the second HTML
        string disabledHtml = File.ReadAllText(disabledHtmlPath);

        // Compare the two HTML strings
        bool areIdentical = string.Equals(defaultHtml, disabledHtml, StringComparison.Ordinal);
        Console.WriteLine($"HTML outputs are {(areIdentical ? "identical" : "different")}.");

        // Show a short snippet if they differ
        if (!areIdentical)
        {
            Console.WriteLine("\n--- Default HTML snippet ---");
            Console.WriteLine(defaultHtml.Substring(0, Math.Min(200, defaultHtml.Length)));
            Console.WriteLine("\n--- Disabled comments HTML snippet ---");
            Console.WriteLine(disabledHtml.Substring(0, Math.Min(200, disabledHtml.Length)));
        }
    }
}