using System;
using System.IO;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsHtmlLineBreakDemo
{
    class Program
    {
        static void Main()
        {
            // Sample HTML containing <br> line break tags
            string html = "<p>First line<br>Second line<br/>Third line</p>";

            // Configure HTML load options (keep spaces as they are)
            HtmlLoadOptions loadOptions = new HtmlLoadOptions(LoadFormat.Html);
            loadOptions.DeleteRedundantSpaces = false; // preserve spaces around line breaks

            // Load the HTML into a workbook using a memory stream
            byte[] htmlBytes = Encoding.UTF8.GetBytes(html);
            using (MemoryStream stream = new MemoryStream(htmlBytes))
            {
                Workbook workbook = new Workbook(stream, loadOptions);
                Worksheet worksheet = workbook.Worksheets[0];
                Cell cell = worksheet.Cells["A1"];

                // Retrieve the cell's text after import
                string cellText = cell.StringValue;

                // Verify that line break tags have been converted to newline characters
                bool containsNewLine = cellText.Contains("\n") || cellText.Contains("\r");

                Console.WriteLine("Cell text after HTML import:");
                Console.WriteLine(cellText);
                Console.WriteLine();
                Console.WriteLine("Line break conversion successful: " + containsNewLine);
            }
        }
    }
}