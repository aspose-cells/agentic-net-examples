// Title: How to convert an HTML table with colspan and rowspan into merged cells in an Excel worksheet using Aspose.Cells for .NET
// AI Prompts: Load an HTML string that includes colspan and rowspan attributes into a Workbook with HtmlLoadOptions and save it as an XLSX file while preserving merged cells. | Use a MemoryStream to feed HTML markup to Aspose.Cells and generate an Excel workbook that reflects the original table's merged cells in C#. | Show the steps to convert HTML tables with row and column spans into Excel merged cells by loading the HTML directly into a Workbook and exporting to .xlsx.
// Common Searches: aspnet convert html table with colspan and rowspan to excel merged cells using aspose.cells | c# load html string with merged cells into workbook preserving row span | aspose.cells htmlloadoptions keep merged cells from html table | how to export html table with merged cells to xlsx in .net | memorystream html to workbook aspose.cells merged cells example
// Tags: html to excel merged cells Aspose.Cells | HtmlLoadOptions preserve colspan rowspan | C# memorystream html workbook conversion | export html table as xlsx with merged cells | Aspose.Cells load html with merged cells

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsHtmlConversion
{
    // The sample loads an HTML string containing a table with colspan and rowspan attributes into an Aspose.Cells Workbook using HtmlLoadOptions and a MemoryStream, then saves the workbook as an XLSX file. The resulting Excel worksheet retains the merged cells defined in the original HTML table.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // HTML content containing tables with colspan and rowspan.
                string html = @"
<table border='1'>
    <tr>
        <th>Header 1</th>
        <th colspan='2'>Header 2-3</th>
        <th>Header 4</th>
    </tr>
    <tr>
        <td rowspan='2'>Row 1-2, Col 1</td>
        <td>R1C2</td>
        <td>R1C3</td>
        <td>R1C4</td>
    </tr>
    <tr>
        <td colspan='2'>R2C2-3 (merged)</td>
        <td>R2C4</td>
    </tr>
</table>";

                // Load the HTML into a new workbook using HtmlLoadOptions.
                var loadOptions = new HtmlLoadOptions();
                Workbook workbook;

                // Convert HTML string to a memory stream.
                byte[] htmlBytes = Encoding.UTF8.GetBytes(html);
                using (var ms = new MemoryStream(htmlBytes))
                {
                    workbook = new Workbook(ms, loadOptions);
                }

                // Define output file path.
                string outputPath = "ConvertedTable.xlsx";

                // Ensure the directory for the output file exists.
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook to an Excel file.
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred during conversion:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
