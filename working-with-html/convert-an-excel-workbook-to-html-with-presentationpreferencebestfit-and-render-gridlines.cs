// Title: Export Excel to HTML with Best‑Fit Layout and Gridlines using Aspose.Cells for .NET
// Description: Shows how to load an .xlsx workbook, configure HtmlSaveOptions with PresentationPreference (best‑fit) and ExportGridLines, and save the file as HTML, preserving column widths and displaying gridlines for a web‑ready view.
// Keywords: Aspose.Cells | C# HTML export | PresentationPreference | best fit HTML | ExportGridLines | Excel to HTML conversion | preserve column width | gridlines in HTML | .NET Aspose.Cells example
// Common Searches: Aspose.Cells export Excel to HTML best fit | Enable gridlines when saving workbook as HTML with Aspose.Cells | C# HtmlSaveOptions PresentationPreference example | Convert .xlsx to HTML with column width preservation | Aspose.Cells HTMLSaveOptions ExportGridLines property
// Developer Intent: Create an HTML representation of an Excel workbook that retains the original column sizing (best‑fit) and shows gridlines, using Aspose.Cells for .NET.
// Use Cases: Generate printable web reports that keep Excel column widths and display gridlines for clear data separation. | Embed Excel data in web portals or dashboards where gridlines improve readability. | Automate bulk conversion of multiple .xlsx files to HTML with consistent best‑fit styling and visible gridlines. | Provide client‑side previews of spreadsheets without requiring Microsoft Office.
// AI Prompts: Show C# code that loads an Excel file, sets HtmlSaveOptions.PresentationPreference = true and ExportGridLines = true, then saves as HTML with Aspose.Cells. | Explain how PresentationPreference influences column width and layout when exporting an Excel workbook to HTML using Aspose.Cells. | Write a PowerShell script that invokes a compiled .NET assembly to convert an .xlsx file to HTML with best‑fit rendering and gridlines enabled. | Give a step‑by‑step guide to batch convert multiple Excel files to HTML while preserving gridlines using Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    // Shows how to load an .xlsx workbook, configure HtmlSaveOptions with PresentationPreference (best‑fit) and ExportGridLines, and save the file as HTML, preserving column widths and displaying gridlines for a web‑ready view.
    class Program
    {
        static void Main(string[] args)
        {
            // Load an existing Excel workbook from file
            // Replace "input.xlsx" with the path to your source workbook
            Workbook workbook = new Workbook("input.xlsx");

            // Create HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

            // Enable presentation preference (best‑fit rendering)
            htmlOptions.PresentationPreference = true;

            // Export gridlines so they appear in the generated HTML
            htmlOptions.ExportGridLines = true;

            // Save the workbook as an HTML file using the configured options
            // Replace "output.html" with the desired output path
            workbook.Save("output.html", htmlOptions);

            Console.WriteLine("Workbook has been successfully converted to HTML with best‑fit presentation and gridlines.");
        }
    }
}
