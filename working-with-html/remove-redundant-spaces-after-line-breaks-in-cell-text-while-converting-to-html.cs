using System;
using System.IO;
using System.Text;
using Aspose.Cells;

class RemoveRedundantSpacesDemo
{
    static void Main()
    {
        // Sample HTML containing redundant spaces after line breaks
        string html = "<p>Line1<br>   Line2   <br>   Line3</p>";

        // Convert the HTML string to a memory stream
        byte[] htmlBytes = Encoding.UTF8.GetBytes(html);
        using (MemoryStream htmlStream = new MemoryStream(htmlBytes))
        {
            // Load the workbook from the HTML stream with redundant spaces removal enabled
            HtmlLoadOptions loadOptions = new HtmlLoadOptions();
            loadOptions.DeleteRedundantSpaces = true; // removes extra spaces after <br>

            Workbook workbook = new Workbook(htmlStream, loadOptions);
            Worksheet worksheet = workbook.Worksheets[0];

            // Optional: display the cleaned cell text
            Console.WriteLine("Cell A1 after loading: '" + worksheet.Cells["A1"].StringValue + "'");

            // Save the workbook back to HTML (spaces are already cleaned)
            HtmlSaveOptions saveOptions = new HtmlSaveOptions();
            workbook.Save("output.html", saveOptions);
        }
    }
}