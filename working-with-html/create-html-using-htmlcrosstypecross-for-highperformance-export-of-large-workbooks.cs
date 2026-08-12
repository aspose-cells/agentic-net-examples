// Title: High‑Performance HTML Export of Large Workbooks with HtmlCrossType.Cross (Aspose.Cells for .NET)
// Description: Demonstrates creating a workbook with 5,000 rows, configuring HtmlSaveOptions to use HtmlCrossType.Cross, optionally disabling CSS, and saving the data as a single HTML file for fast, low‑memory conversion.
// Keywords: Aspose.Cells | HtmlCrossType.Cross | HTML export | large workbook | .NET | C# | performance | HtmlSaveOptions | disable CSS
// Common Searches: Aspose.Cells HtmlCrossType.Cross example | export large Excel to HTML .NET | speed up HTML conversion Aspose.Cells | disable CSS in HtmlSaveOptions | high performance HTML export C#
// Developer Intent: Generate a single HTML file from a massive workbook quickly and with minimal memory usage by using HtmlCrossType.Cross.
// Use Cases: Render extensive data sets on a web page without loading the full Excel file. | Accelerate reporting pipelines by turning off CSS generation during HTML conversion. | Provide on‑the‑fly HTML previews in ASP.NET Core APIs for large spreadsheets.
// AI Prompts: Write C# code that creates a 10,000‑row workbook and exports it to HTML using HtmlCrossType.Cross with Aspose.Cells. | Explain how HtmlCrossStringType.Cross reduces memory consumption and improves export speed in Aspose.Cells. | Show how to stream the HTML output directly to an ASP.NET Core response using HtmlSaveOptions.

using System;
using Aspose.Cells;

namespace AsposeCellsHtmlCrossDemo
{
    // Demonstrates creating a workbook with 5,000 rows, configuring HtmlSaveOptions to use HtmlCrossType.Cross, optionally disabling CSS, and saving the data as a single HTML file for fast, low‑memory conversion.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate the worksheet with a large amount of sample data
            // (In a real scenario this could be loaded from a source or generated)
            for (int row = 0; row < 5000; row++)
            {
                for (int col = 0; col < 20; col++)
                {
                    sheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // Create HTML save options (lifecycle rule)
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

            // Set the cross‑cell string handling to Cross for high‑performance export
            // (uses HtmlSaveOptions.HtmlCrossStringType property)
            htmlOptions.HtmlCrossStringType = HtmlCrossType.Cross;

            // Optional: improve performance further by disabling CSS (if desired)
            // htmlOptions.DisableCss = true;

            // Save the workbook as HTML using the configured options (lifecycle rule)
            workbook.Save("LargeWorkbook.html", htmlOptions);

            Console.WriteLine("HTML export completed with HtmlCrossType.Cross.");
        }
    }
}
