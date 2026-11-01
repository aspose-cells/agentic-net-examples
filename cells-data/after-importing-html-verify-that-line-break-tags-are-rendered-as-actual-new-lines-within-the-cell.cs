// Title: Load HTML with <br> tags into an Aspose.Cells workbook and confirm they appear as new line characters in the cell (C#)
// AI Prompts: Use HtmlLoadOptions to import an HTML string that contains <br> elements into a Workbook and return the text of cell A1. | Check the retrieved cell value for newline characters (\n or \r) to verify that the <br> tags were converted to line breaks.
// Common Searches: Aspose.Cells C# load HTML and keep <br> tags as line breaks in Excel cells | How to detect newline characters in a cell after importing HTML with HtmlLoadOptions | Preserve line breaks from HTML when converting to workbook using Aspose.Cells | Verify that <br> tags become actual new lines in a cell after HTML import in C#
// Tags: html to workbook line break conversion | Aspose.Cells HtmlLoadOptions preserve newlines | detect newline in cell stringvalue c# | import html br as newline excel cell | verify line break rendering aspose cells

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

// The example loads an HTML string containing a <br> tag into a Workbook via HtmlLoadOptions, accesses cell A1, retrieves its StringValue, checks for newline characters, and outputs whether the line break was rendered correctly.
class Program
{
    static void Main()
    {
        try
        {
            // HTML containing a line break tag (<br>)
            string html = "<html><body><table><tr><td>Line1<br>Line2</td></tr></table></body></html>";

            // Convert HTML string to a memory stream
            byte[] htmlBytes = Encoding.UTF8.GetBytes(html);
            using (MemoryStream htmlStream = new MemoryStream(htmlBytes))
            {
                // Load the HTML into a new workbook using HtmlLoadOptions
                HtmlLoadOptions loadOptions = new HtmlLoadOptions();
                Workbook workbook = new Workbook(htmlStream, loadOptions);

                // Get the first worksheet and the first cell (A1) where the HTML content is placed
                Worksheet sheet = workbook.Worksheets[0];
                Cell cell = sheet.Cells[0, 0]; // A1

                // Retrieve the cell's displayed text
                string cellText = cell.StringValue;

                // Verify that the <br> tag was rendered as an actual new line character
                bool hasNewLine = cellText.Contains("\n") || cellText.Contains("\r");

                // Output the result
                Console.WriteLine("Cell content:");
                Console.WriteLine(cellText);
                Console.WriteLine("Line break rendered as new line: " + hasNewLine);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
