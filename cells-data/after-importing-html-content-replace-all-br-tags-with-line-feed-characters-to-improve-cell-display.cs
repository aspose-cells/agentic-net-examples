// Title: C# – Replace <br> Tags with Newlines After Importing HTML into Aspose.Cells Workbook
// Description: Load HTML containing <br> tags into a Workbook via MemoryStream and HtmlLoadOptions, then use Workbook.Replace("<br>", "\n") to convert each tag to a line feed before saving as XLSX.
// Keywords: Aspose.Cells HTML import | C# replace br tag | convert <br> to newline | Workbook.Replace method | Excel line break handling
// Common Searches: Aspose.Cells replace <br> with \n C# | how to convert HTML line breaks to Excel newlines | C# load HTML into workbook and preserve line breaks | Aspose.Cells HTMLLoadOptions line break replacement
// Developer Intent: Swap every <br> tag in imported HTML for a newline character so text appears on separate lines inside Excel cells.
// Use Cases: Import an HTML snippet with <br> tags and display each segment on a new line in a single cell. | Transform an HTML report that uses <br> for paragraph breaks into a properly formatted XLSX file. | Process HTML email bodies, replace their line‑break tags, and generate an Excel worksheet with readable text.
// AI Prompts: Show C# code that loads HTML into an Aspose.Cells workbook and replaces all <br> tags with "\n" before saving. | Explain how to use HtmlLoadOptions together with Workbook.Replace to handle HTML line breaks in Aspose.Cells. | Provide a robust method for replacing both <br> and <br/> tags with newline characters in a C# Aspose.Cells project.

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsHtmlImport
{
    // Load HTML containing <br> tags into a Workbook via MemoryStream and HtmlLoadOptions, then use Workbook.Replace("<br>", "\n") to convert each tag to a line feed before saving as XLSX.
    class Program
    {
        static void Main()
        {
            // Sample HTML content containing <br> tags
            string html = "<p>Hello<br>World<br>From Aspose.Cells</p>";

            // Convert the HTML string to a memory stream (required for Workbook loading)
            byte[] htmlBytes = Encoding.UTF8.GetBytes(html);
            using (MemoryStream stream = new MemoryStream(htmlBytes))
            {
                // Set HTML load options – enable deletion of redundant spaces (optional)
                HtmlLoadOptions loadOptions = new HtmlLoadOptions();
                loadOptions.DeleteRedundantSpaces = true;

                // Load the HTML content into a workbook
                Workbook workbook = new Workbook(stream, loadOptions);

                // Replace all <br> tags with line feed characters to improve cell display
                workbook.Replace("<br>", "\n");

                // Save the resulting workbook
                workbook.Save("Output.xlsx");
            }

            Console.WriteLine("HTML imported and <br> tags replaced. Workbook saved as Output.xlsx");
        }
    }
}
