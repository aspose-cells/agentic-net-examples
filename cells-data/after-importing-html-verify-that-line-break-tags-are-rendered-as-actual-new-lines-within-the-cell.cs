// Title: C# Aspose.Cells: Import HTML with <br> tags and confirm line breaks in a worksheet cell
// Description: This example shows how to load an HTML string that contains <br> tags into a Workbook using a MemoryStream and HtmlLoadOptions (DeleteRedundantSpaces = false). It reads the text of cell A1, checks for newline characters (\n or \r), prints the result, and optionally saves the file as an XLSX document.
// Keywords: Aspose.Cells HTML import C# | preserve line breaks from <br> tags | HtmlLoadOptions DeleteRedundantSpaces | verify newline characters in Excel cell | load HTML from memory stream Aspose.Cells | C# Excel line break handling
// Common Searches: Aspose.Cells keep <br> as line break | C# load HTML into Excel workbook with line breaks | HtmlLoadOptions preserve newline characters | check cell text for \n after HTML import | convert HTML paragraph to multi‑line Excel cell
// Developer Intent: Load an HTML fragment that uses <br> tags into an Excel worksheet and ensure the tags are translated into actual line‑break characters inside the target cell.
// Use Cases: Transform an HTML email body into an Excel report while retaining its original line formatting. | Import web‑scraped table data where cells contain <br> separators and need to appear as multi‑line entries. | Automated validation that exported Excel files preserve the line‑break structure of source HTML content.
// AI Prompts: Generate C# code that configures HtmlLoadOptions to keep <br> tags as new lines when loading HTML into Aspose.Cells. | Write a unit test that asserts cell A1 contains a newline after importing HTML with line‑break tags using Aspose.Cells. | Explain how to replace <br> tags with "\n" before loading HTML if Aspose.Cells does not automatically preserve them.

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsHtmlLineBreakDemo
{
    // This example shows how to load an HTML string that contains <br> tags into a Workbook using a MemoryStream and HtmlLoadOptions (DeleteRedundantSpaces = false). It reads the text of cell A1, checks for newline characters (\n or \r), prints the result, and optionally saves the file as an XLSX document.
    class Program
    {
        static void Main()
        {
            // HTML containing <br> tags that should become line breaks in the cell
            string htmlContent = "<p>First line<br>Second line<br/>Third line</p>";

            // Convert the HTML string to a UTF‑8 memory stream
            byte[] htmlBytes = Encoding.UTF8.GetBytes(htmlContent);
            using (MemoryStream htmlStream = new MemoryStream(htmlBytes))
            {
                // Configure HTML load options (optional: keep default DeleteRedundantSpaces)
                HtmlLoadOptions loadOptions = new HtmlLoadOptions(LoadFormat.Html);
                loadOptions.DeleteRedundantSpaces = false; // preserve spaces around <br> tags

                // Load the HTML into a workbook
                Workbook workbook = new Workbook(htmlStream, loadOptions);
                Worksheet sheet = workbook.Worksheets[0];

                // Retrieve the value of the first cell (A1) after import
                string cellText = sheet.Cells["A1"].StringValue;

                // Output the raw cell text to the console
                Console.WriteLine("Cell A1 text after HTML import:");
                Console.WriteLine(cellText);

                // Verify that line break characters are present
                bool containsLineBreak = cellText.Contains("\n") || cellText.Contains("\r");
                Console.WriteLine("Contains line break characters: " + containsLineBreak);

                // Save the workbook for further inspection (optional)
                workbook.Save("ImportedHtml.xlsx", SaveFormat.Xlsx);
            }
        }
    }
}
