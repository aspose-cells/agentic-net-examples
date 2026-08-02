// Title: Hide Zero Values in Excel and Export to HTML with Aspose.Cells for .NET
// Description: Loads an existing workbook, turns off zero‑value display on every worksheet, and saves the result as an HTML file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | DisplayZeros | suppress zeroes | Excel to HTML conversion | HtmlSaveOptions | Workbook.Save | export workbook as HTML | Excel reporting | web‑ready spreadsheet
// Common Searches: Aspose.Cells hide zero values when converting to HTML | C# set DisplayZeros false for all worksheets | Export Excel workbook to HTML without zeroes Aspose | Disable zero display in Aspose.Cells HTML export | Convert .xlsx to .html using Aspose.Cells and omit zeros
// Developer Intent: Load an Excel file, disable zero‑value rendering for each sheet, and generate an HTML version of the workbook.
// Use Cases: Produce web‑friendly financial reports that exclude zero balances. | Create clean HTML snapshots of dashboards without placeholder zeros. | Automate batch conversion of multiple spreadsheets to HTML while suppressing empty cells.
// AI Prompts: Generate C# code that opens a workbook, sets Worksheet.DisplayZeros = false for every sheet, and saves it as HTML with optional CSS embedding using Aspose.Cells. | Explain how HtmlSaveOptions can be configured to control image handling, CSS inclusion, and other settings when exporting Excel to HTML with Aspose.Cells. | Write a script that processes a folder of .xlsx files, hides zero values, and outputs corresponding .html files using Aspose.Cells for .NET.

using System;
using Aspose.Cells;

// Loads an existing workbook, turns off zero‑value display on every worksheet, and saves the result as an HTML file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Load the existing Excel workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Disable displaying zero values for all worksheets
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            sheet.DisplayZeros = false;
        }

        // Create HTML save options (default settings)
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

        // Export the workbook to an HTML page
        workbook.Save("output.html", htmlOptions);
    }
}
