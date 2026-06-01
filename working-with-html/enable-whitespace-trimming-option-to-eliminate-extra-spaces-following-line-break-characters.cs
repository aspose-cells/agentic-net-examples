using System;
using System.IO;
using System.Text;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Sample HTML containing line breaks (<br>) and redundant spaces after them
        string html = "<p>Line1<br>   Line2<br>   Line3</p>";

        // Create HtmlLoadOptions and enable whitespace trimming
        HtmlLoadOptions loadOptions = new HtmlLoadOptions();
        loadOptions.DeleteRedundantSpaces = true; // eliminates extra spaces after <br> tags

        // Convert the HTML string to a memory stream for loading
        byte[] htmlBytes = Encoding.UTF8.GetBytes(html);
        using (MemoryStream stream = new MemoryStream(htmlBytes))
        {
            // Load the workbook using the specified options
            Workbook workbook = new Workbook(stream, loadOptions);

            // Save the resulting workbook (the cells will contain trimmed text)
            workbook.Save("TrimmedOutput.xlsx");
        }
    }
}