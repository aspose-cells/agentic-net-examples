using System;
using System.IO;
using System.Text;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Sample HTML containing <br> tags
        string html = "<p>First line<br>Second line<br/>Third line</p>";

        // Convert the HTML string to a memory stream
        byte[] htmlBytes = Encoding.UTF8.GetBytes(html);
        using (MemoryStream stream = new MemoryStream(htmlBytes))
        {
            // Load the HTML into a workbook with HtmlLoadOptions
            HtmlLoadOptions loadOptions = new HtmlLoadOptions();
            // Optional: delete redundant spaces that may appear after line breaks
            loadOptions.DeleteRedundantSpaces = true;

            Workbook workbook = new Workbook(stream, loadOptions);

            // Replace <br> and <br/> tags with line feed characters in all cells
            workbook.Replace("<br>", "\n");
            workbook.Replace("<br/>", "\n");

            // Save the resulting workbook
            workbook.Save("output.xlsx");
        }
    }
}