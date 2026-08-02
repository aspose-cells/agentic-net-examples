using System;
using System.IO;
using System.Text;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Sample HTML containing <br> tags with extra spaces after line breaks
        string html = "<p>Line1<br>   Line2<br>    Line3</p>";

        // Enable whitespace trimming after line breaks
        HtmlLoadOptions loadOptions = new HtmlLoadOptions
        {
            DeleteRedundantSpaces = true
        };

        // Load the HTML into a workbook using the configured options
        byte[] htmlBytes = Encoding.UTF8.GetBytes(html);
        using (MemoryStream stream = new MemoryStream(htmlBytes))
        {
            Workbook workbook = new Workbook(stream, loadOptions);

            // Demonstrate that redundant spaces have been removed
            Console.WriteLine("Cell A1 content after trimming: " +
                workbook.Worksheets[0].Cells["A1"].StringValue);

            // Save the workbook (optional)
            workbook.Save("TrimmedOutput.xlsx");
        }
    }
}

// Author: Aspose.Cells .NET example