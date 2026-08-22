// Title: Check that <br> tags in imported HTML are rendered as line breaks in an Aspose.Cells C# worksheet cell
// AI Prompts: Load an HTML snippet containing <br> elements into a Workbook using HtmlLoadOptions and read the resulting cell value. | Convert the HTML string to a MemoryStream, import it with Aspose.Cells, and programmatically detect newline characters in the cell text. | Save the workbook after the HTML import and output a boolean indicating whether the cell contains line‑break characters.
// Common Searches: Aspose.Cells C# import HTML with <br> tags and keep line breaks | How to detect newline characters after loading HTML into an Excel cell using Aspose.Cells | C# example for converting HTML <br> tags to Excel cell line breaks with HtmlLoadOptions | Verify line break rendering when loading HTML into Aspose.Cells workbook
// Tags: html import newline handling Aspose.Cells C# | HtmlLoadOptions line break conversion | memory stream html to workbook Aspose.Cells | cell value includes carriage return after HTML load | preserve <br> tags in Excel import Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsHtmlLineBreakDemo
{
    // The example builds an HTML string with <br> tags, loads it into an Aspose.Cells Workbook via HtmlLoadOptions using a MemoryStream, reads the cell's string value to confirm newline characters are present, prints the verification result, and optionally saves the workbook as an XLSX file.
    class Program
    {
        static void Main()
        {
            // Sample HTML containing <br> tags
            string htmlContent = "<p>First line<br>Second line<br/>Third line</p>";

            // Convert the HTML string to a memory stream
            byte[] htmlBytes = System.Text.Encoding.UTF8.GetBytes(htmlContent);
            using (MemoryStream htmlStream = new MemoryStream(htmlBytes))
            {
                // Load the HTML into a workbook using HtmlLoadOptions
                HtmlLoadOptions loadOptions = new HtmlLoadOptions(LoadFormat.Html);
                // DeleteRedundantSpaces is false by default; keep it unchanged
                Workbook workbook = new Workbook(htmlStream, loadOptions);

                // Access the first worksheet and the cell where the HTML was imported (A1)
                Worksheet worksheet = workbook.Worksheets[0];
                Cell cell = worksheet.Cells["A1"];

                // Retrieve the cell's string value; Aspose.Cells converts <br> to newline characters
                string cellText = cell.StringValue;

                // Output the cell content to verify line breaks
                Console.WriteLine("Cell A1 content after HTML import:");
                Console.WriteLine(cellText);

                // Check whether the text contains newline characters
                bool containsNewLine = cellText.Contains("\n") || cellText.Contains("\r");
                Console.WriteLine("Line break rendered as new line: " + containsNewLine);

                // Save the workbook to an Excel file for further inspection (optional)
                workbook.Save("HtmlImportResult.xlsx", SaveFormat.Xlsx);
            }
        }
    }
}
