// Title: Convert Excel to HTML with Gridlines Using Aspose.Cells for .NET (C#)
// Description: Loads an .xlsx workbook, applies default HtmlSaveOptions, turns on ExportGridLines, and saves the file as HTML so the worksheet gridlines appear in the output.
// Keywords: Aspose.Cells | C# | .NET | Excel to HTML | HtmlSaveOptions | ExportGridLines | gridlines | Workbook.Save | HTML export | convert xlsx to html
// Common Searches: Aspose.Cells export Excel to HTML with gridlines | C# convert xlsx to html preserving gridlines | HtmlSaveOptions ExportGridLines example | How to save Excel as HTML using Aspose.Cells .NET | default HTML export options Aspose.Cells
// Developer Intent: Generate an HTML representation of an Excel workbook that retains the original gridlines.
// Use Cases: Display spreadsheet data on a web portal while keeping the familiar grid layout. | Provide a lightweight HTML download for users who do not have Excel installed. | Integrate Excel‑to‑HTML conversion into a .NET reporting service that requires visible cell borders.
// AI Prompts: Write C# code with Aspose.Cells to convert an .xlsx file to HTML and enable gridlines using default options. | Explain the effect of HtmlSaveOptions.ExportGridLines on the produced HTML and how to combine it with other export settings. | Show how to export each worksheet of a workbook to separate HTML files while preserving gridlines.

using System;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    // Loads an .xlsx workbook, applies default HtmlSaveOptions, turns on ExportGridLines, and saves the file as HTML so the worksheet gridlines appear in the output.
    class Program
    {
        static void Main()
        {
            // Load an existing Excel workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Create HTML save options with default settings
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

            // Enable exporting of gridlines
            htmlOptions.ExportGridLines = true;

            // Save the workbook as an HTML file using the specified options
            workbook.Save("output.html", htmlOptions);
        }
    }
}
