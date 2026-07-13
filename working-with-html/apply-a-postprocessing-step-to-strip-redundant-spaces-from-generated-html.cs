using System;
using System.IO;
using System.Text;
using Aspose.Cells;

// Author: Aspose.Cells .NET example – loads HTML, removes redundant spaces, and saves cleaned HTML.
class Program
{
    static void Main()
    {
        // Sample HTML containing redundant spaces
        string html = "<p>   This    text   has   redundant   spaces   </p>";

        // Configure load options to delete redundant spaces during import
        HtmlLoadOptions loadOptions = new HtmlLoadOptions
        {
            DeleteRedundantSpaces = true
        };

        // Convert the HTML string to a memory stream for loading
        byte[] htmlBytes = Encoding.UTF8.GetBytes(html);
        using (MemoryStream stream = new MemoryStream(htmlBytes))
        {
            // Load the HTML into a workbook using the specified options
            Workbook workbook = new Workbook(stream, loadOptions);

            // Save the workbook back to HTML; the output will have redundant spaces removed
            HtmlSaveOptions saveOptions = new HtmlSaveOptions();
            workbook.Save("output.html", saveOptions);
        }
    }
}