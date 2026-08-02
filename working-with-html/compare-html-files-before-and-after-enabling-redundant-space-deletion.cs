using System;
using System.IO;
using System.Text;
using Aspose.Cells;

// Author: Aspose.Cells .NET example – compare HTML loading with and without redundant space deletion
class Program
{
    static void Main()
    {
        // Sample HTML containing redundant spaces
        string html = "<p>   This    text   has   redundant   spaces   </p>";

        // Load the HTML without deleting redundant spaces (default behavior)
        Workbook wbOriginal = LoadHtmlToWorkbook(html, deleteRedundantSpaces: false);

        // Load the same HTML with redundant spaces removed
        Workbook wbTrimmed = LoadHtmlToWorkbook(html, deleteRedundantSpaces: true);

        // Retrieve the text from the first cell of each workbook
        string originalText = wbOriginal.Worksheets[0].Cells["A1"].StringValue;
        string trimmedText = wbTrimmed.Worksheets[0].Cells["A1"].StringValue;

        // Output the results
        Console.WriteLine("Original cell text: \"" + originalText + "\"");
        Console.WriteLine("Trimmed cell text:  \"" + trimmedText + "\"");

        // Simple comparison to show the effect of DeleteRedundantSpaces
        if (originalText == trimmedText)
            Console.WriteLine("No difference detected.");
        else
            Console.WriteLine("Difference detected.");
    }

    // Helper method to load HTML into a Workbook with the specified DeleteRedundantSpaces setting
    static Workbook LoadHtmlToWorkbook(string htmlContent, bool deleteRedundantSpaces)
    {
        HtmlLoadOptions loadOptions = new HtmlLoadOptions
        {
            DeleteRedundantSpaces = deleteRedundantSpaces
        };

        byte[] htmlBytes = Encoding.UTF8.GetBytes(htmlContent);
        using (MemoryStream stream = new MemoryStream(htmlBytes))
        {
            // Load the HTML stream into a workbook using the provided options
            return new Workbook(stream, loadOptions);
        }
    }
}