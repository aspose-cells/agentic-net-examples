// Title: Export Excel to HTML without Zero Values using Aspose.Cells for .NET (C#)
// Description: Loads an existing workbook, disables zero‑value display on every worksheet (DisplayZeros = false), and saves the file as an HTML page with default HtmlSaveOptions. The resulting HTML omits cells that contain zero.
// Keywords: Aspose.Cells C# | .NET Excel to HTML | Hide zero values Aspose | DisplayZeros false | ShowZeroValues property | Workbook.Save HTML | HtmlSaveOptions Aspose | Excel web export | Suppress zero cells | Convert XLSX to HTML
// Common Searches: Aspose.Cells hide zero values when exporting to HTML | C# export Excel workbook to HTML without zeros | DisplayZeros property Aspose.Cells example | Convert XLSX to HTML using Aspose.Cells .NET | How to suppress zero cells in HTML output from Excel
// Developer Intent: Load an existing Excel file, turn off zero‑value display for all worksheets, and generate an HTML version of the workbook.
// Use Cases: Create web‑ready financial reports that exclude zero entries. | Produce clean printable HTML snapshots of spreadsheets for dashboards. | Automate batch conversion of multiple workbooks to HTML while removing zero clutter.
// AI Prompts: Generate C# code with Aspose.Cells that loads a workbook, sets DisplayZeros = false for each worksheet, and saves it as HTML. | Write a reusable method that accepts input and output paths and converts an Excel file to HTML with zero values hidden using Aspose.Cells. | Explain how to customize HtmlSaveOptions (e.g., embed CSS, inline images) when exporting Excel to HTML while suppressing zero cells.

using System;
using Aspose.Cells;

// Loads an existing workbook, disables zero‑value display on every worksheet (DisplayZeros = false), and saves the file as an HTML page with default HtmlSaveOptions. The resulting HTML omits cells that contain zero.
class ExportExcelToHtml
{
    static void Main()
    {
        // Load the existing Excel workbook from file
        Workbook workbook = new Workbook("input.xlsx");

        // Disable displaying zero values for all worksheets
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            sheet.DisplayZeros = false;
        }

        // Create HTML save options (default settings)
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

        // Export the entire workbook to an HTML file
        workbook.Save("output.html", htmlOptions);
    }
}
