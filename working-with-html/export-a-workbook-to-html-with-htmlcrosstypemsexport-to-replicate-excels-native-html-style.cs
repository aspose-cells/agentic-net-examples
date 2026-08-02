// Title: Export Excel to HTML with Aspose.Cells using HtmlCrossType.MSExport (C#)
// Description: Shows how to build a workbook, populate cells, configure HtmlSaveOptions.HtmlCrossStringType to HtmlCrossType.MSExport, and save the file as HTML that mirrors Excel's built‑in HTML export.
// Keywords: Aspose.Cells | HtmlCrossType | MSExport | C# | HTML export | Excel to HTML | cross‑cell rendering | HtmlSaveOptions | .NET | workbook to HTML
// Common Searches: Aspose.Cells HtmlCrossType.MSExport example | C# export Excel workbook to HTML with Excel style | HtmlSaveOptions cross‑cell string rendering | How to keep merged cells when converting Excel to HTML | Generate Excel‑like HTML using Aspose.Cells
// Developer Intent: Convert an Excel workbook to HTML while preserving Excel’s native formatting (merged cells, cross‑cell strings) by setting HtmlCrossType.MSExport in Aspose.Cells for a C# project.
// Use Cases: Create web‑ready reports that look identical to Excel’s native HTML output. | Automate conversion of spreadsheets with merged cells for email previews or documentation. | Build a portal that displays uploaded Excel files as faithful HTML pages without losing layout.
// AI Prompts: Generate C# code that loads an existing .xlsx file and saves it as HTML using HtmlCrossType.MSExport with Aspose.Cells. | Explain how to configure HtmlSaveOptions to embed CSS, images, and enable MSExport cross‑cell rendering for accurate Excel‑style HTML. | Compare HtmlCrossType.MSExport with HtmlCrossType.Default and suggest when each should be used.

using System;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    // Shows how to build a workbook, populate cells, configure HtmlSaveOptions.HtmlCrossStringType to HtmlCrossType.MSExport, and save the file as HTML that mirrors Excel's built‑in HTML export.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook (empty Excel file)
            Workbook workbook = new Workbook();

            // Access the first worksheet to add some sample data
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("First");
            sheet.Cells["B1"].PutValue("Second");
            sheet.Cells["A2"].PutValue(123);
            sheet.Cells["B2"].PutValue(456);

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

            // Set the cross-cell string rendering to mimic Excel's native HTML export
            htmlOptions.HtmlCrossStringType = HtmlCrossType.MSExport;

            // Save the workbook as an HTML file using the configured options
            workbook.Save("output.html", htmlOptions);

            Console.WriteLine("Workbook exported to HTML with HtmlCrossType.MSExport.");
        }
    }
}
