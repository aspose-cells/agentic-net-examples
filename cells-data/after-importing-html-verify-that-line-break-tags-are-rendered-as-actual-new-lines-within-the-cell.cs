using System;
using System.IO;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsHtmlImportDemo
{
    class Program
    {
        static void Main()
        {
            // HTML content with <br> tags representing line breaks
            string htmlContent = "<p>First line<br>Second line<br/>Third line</p>";

            // Convert the HTML string to a UTF-8 byte array and load it via a MemoryStream
            byte[] htmlBytes = Encoding.UTF8.GetBytes(htmlContent);
            using (MemoryStream htmlStream = new MemoryStream(htmlBytes))
            {
                // Configure HTML load options (default settings are sufficient for <br> handling)
                HtmlLoadOptions loadOptions = new HtmlLoadOptions(LoadFormat.Html);
                // Optional: keep spaces as they appear in the HTML
                loadOptions.DeleteRedundantSpaces = false;

                // Load the HTML into a workbook
                Workbook workbook = new Workbook(htmlStream, loadOptions);
                Worksheet sheet = workbook.Worksheets[0];

                // Retrieve the value of the first cell (A1) where the HTML was imported
                Cell cell = sheet.Cells["A1"];
                string cellText = cell.StringValue;

                // Verify that line break tags have been converted to actual new line characters
                bool containsNewLine = cellText.Contains("\n") || cellText.Contains("\r");
                Console.WriteLine("Cell A1 text:");
                Console.WriteLine(cellText);
                Console.WriteLine();
                Console.WriteLine("Line break conversion successful: " + containsNewLine);
            }
        }
    }
}