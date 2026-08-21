// Title: C# – Export Excel Worksheet to HTML with Gridlines Using Aspose.Cells for .NET
// Description: Load an Excel workbook, enable worksheet gridlines, set HtmlSaveOptions.ExportGridLines, and save the active sheet as an HTML file that preserves the original grid layout.
// Keywords: Aspose.Cells C# HTML export | ExportGridLines example | IsGridlinesVisible .NET | Excel to HTML with gridlines | Aspose.Cells HtmlSaveOptions | C# Excel HTML conversion | gridline rendering Aspose | single worksheet HTML export
// Common Searches: Aspose.Cells export gridlines to HTML C# | How to keep Excel gridlines when saving as HTML | HtmlSaveOptions ExportGridLines usage | Save only active worksheet as HTML Aspose | C# code to display Excel gridlines in web page
// Developer Intent: Create an HTML representation of an Excel sheet that shows the original gridlines.
// Use Cases: Embedding a spreadsheet view in a web portal while retaining cell borders. | Generating printable HTML snapshots of reports that need the grid layout. | Providing a lightweight, browser‑friendly preview of Excel data for dashboards.
// AI Prompts: Generate C# code with Aspose.Cells that loads a workbook and saves the first worksheet as HTML with gridlines visible. | Explain the interaction between Worksheet.IsGridlinesVisible and HtmlSaveOptions.ExportGridLines when exporting to HTML. | Outline the steps to export only the active worksheet to HTML while preserving gridlines using Aspose.Cells for .NET.

using System;
using Aspose.Cells;

namespace AsposeCellsGridlinesHtmlDemo
{
    // Load an Excel workbook, enable worksheet gridlines, set HtmlSaveOptions.ExportGridLines, and save the active sheet as an HTML file that preserves the original grid layout.
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your actual file path)
            string inputPath = "input.xlsx";
            Workbook workbook = new Workbook(inputPath);

            // Ensure gridlines are visible in the worksheet
            Worksheet worksheet = workbook.Worksheets[0];
            worksheet.IsGridlinesVisible = true;

            // Configure HTML save options to export gridlines
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                ExportGridLines = true,               // Enable gridline export
                ExportActiveWorksheetOnly = true      // Export only the active sheet (optional)
            };

            // Save the workbook as HTML with gridlines
            string outputPath = "output.html";
            workbook.Save(outputPath, htmlOptions);

            Console.WriteLine($"HTML file saved to '{outputPath}' with gridlines exported.");
        }
    }
}
