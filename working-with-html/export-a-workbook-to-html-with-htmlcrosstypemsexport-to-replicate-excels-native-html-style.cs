// Title: C# – Export Workbook to HTML with Aspose.Cells using HtmlCrossType.MSExport (Excel‑style HTML)
// Description: Demonstrates how to create a workbook, add data, set HtmlSaveOptions.HtmlCrossStringType to HtmlCrossType.MSExport, and save the file as HTML so the output matches Excel's native HTML formatting.
// Keywords: Aspose.Cells | C# | HTML export | HtmlCrossType.MSExport | Excel style HTML | HtmlSaveOptions | Workbook to HTML | sample code | .NET
// Common Searches: Aspose.Cells export workbook to HTML MSExport C# | HtmlCrossStringType MSExport example | How to mimic Excel HTML output with Aspose.Cells | C# save Excel as HTML with original formatting
// Developer Intent: Generate HTML that replicates Excel’s native formatting by exporting a workbook with HtmlCrossType.MSExport.
// Use Cases: Create web‑ready reports that retain Excel cell styles. | Provide spreadsheet previews in web applications without requiring Office. | Automate batch conversion of Excel files to HTML for email or documentation while preserving the original look.
// AI Prompts: Show how to set HtmlCrossStringType to MSExport in Aspose.Cells and save a workbook as HTML using C#. | Give a C# snippet that exports a workbook to HTML with custom CSS while using HtmlCrossType.MSExport. | Explain the differences between HtmlCrossType.MSExport and the default HTML export option in Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    // Demonstrates how to create a workbook, add data, set HtmlSaveOptions.HtmlCrossStringType to HtmlCrossType.MSExport, and save the file as HTML so the output matches Excel's native HTML formatting.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add some sample data
            sheet.Cells["A1"].PutValue("First");
            sheet.Cells["B1"].PutValue("Second");
            sheet.Cells["A2"].PutValue(123);
            sheet.Cells["B2"].PutValue(456);

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
            // Use MSExport cross string type to mimic Excel's native HTML style
            htmlOptions.HtmlCrossStringType = HtmlCrossType.MSExport;

            // Save the workbook as HTML with the specified options
            workbook.Save("ExportedWithMSExport.html", htmlOptions);

            Console.WriteLine("Workbook exported to HTML using HtmlCrossType.MSExport.");
        }
    }
}
