// Title: Remove Redundant Spaces After <br> Tags with Aspose.Cells for .NET
// Description: Demonstrates loading HTML that contains <br> tags followed by spaces, using HtmlLoadOptions.DeleteRedundantSpaces to strip those spaces during workbook creation, and saving clean HTML without extra blanks after line‑break tags.
// Keywords: Aspose.Cells DeleteRedundantSpaces | C# remove spaces after br tag | HTML to Excel trailing spaces | HtmlLoadOptions example | clean HTML output Aspose.Cells | .NET HTML import whitespace removal
// Common Searches: Aspose.Cells delete spaces after <br> | HtmlLoadOptions DeleteRedundantSpaces C# | remove blank spaces after line break Aspose.Cells | convert HTML to workbook without trailing spaces | clean HTML output Aspose.Cells .NET
// Developer Intent: Load HTML into a Workbook while automatically trimming spaces that follow <br> tags, then save the workbook back to HTML with no redundant whitespace.
// Use Cases: Import an HTML fragment containing <br> tags with trailing spaces and ensure cell values are whitespace‑free. | Export a workbook to HTML and guarantee that generated <br> elements are not followed by extra spaces. | Batch‑process multiple HTML files, applying DeleteRedundantSpaces before converting each to Excel.
// AI Prompts: Write C# code that uses Aspose.Cells HtmlLoadOptions.DeleteRedundantSpaces to clean spaces after <br> tags in an HTML string and saves the result. | Explain the effect of HtmlLoadOptions.DeleteRedundantSpaces when loading HTML into a Workbook and how it influences the saved HTML. | Provide a step‑by‑step guide for batch‑processing HTML files to remove redundant spaces after line‑break tags using Aspose.Cells.

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsHtmlRedundantSpacesDemo
{
    // Demonstrates loading HTML that contains <br> tags followed by spaces, using HtmlLoadOptions.DeleteRedundantSpaces to strip those spaces during workbook creation, and saving clean HTML without extra blanks after line‑break tags.
    class Program
    {
        static void Main()
        {
            // Sample HTML containing <br> tags followed by redundant spaces
            string htmlContent = "<p>Line1<br>   </p><p>Line2<br>    </p>";

            // Prepare load options to delete redundant spaces after <br> tags
            HtmlLoadOptions loadOptions = new HtmlLoadOptions
            {
                DeleteRedundantSpaces = true   // Removes spaces that appear after line‑break tags
            };

            // Convert the HTML string to a memory stream
            byte[] htmlBytes = Encoding.UTF8.GetBytes(htmlContent);
            using (MemoryStream htmlStream = new MemoryStream(htmlBytes))
            {
                // Load the HTML into a workbook using the configured load options
                Workbook workbook = new Workbook(htmlStream, loadOptions);

                // (Optional) Verify that the cell values no longer contain trailing spaces
                // Console.WriteLine($"A1: '{workbook.Worksheets[0].Cells["A1"].StringValue}'");
                // Console.WriteLine($"A2: '{workbook.Worksheets[0].Cells["A2"].StringValue}'");

                // Save the workbook back to HTML; no extra spaces will be present after <br>
                HtmlSaveOptions saveOptions = new HtmlSaveOptions(); // default options are sufficient
                workbook.Save("CleanOutput.html", saveOptions);
            }

            Console.WriteLine("HTML saved to CleanOutput.html without redundant spaces after <br> tags.");
        }
    }
}
