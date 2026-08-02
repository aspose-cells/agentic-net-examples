// Title: Hide Columns C & D and Export Workbook to HTML with Hidden Columns Preserved – Aspose.Cells for .NET
// Description: Loads an existing workbook, hides columns C and D using Cells.HideColumns, configures HtmlSaveOptions with HiddenColDisplayType.Hidden, and saves the file as HTML while keeping the hidden columns in the output.
// Keywords: Aspose.Cells hide columns C D | C# export workbook to HTML | HtmlSaveOptions HiddenColDisplayType | preserve hidden columns Aspose | Aspose.Cells column visibility HTML export
// Common Searches: Aspose.Cells hide specific columns and export to HTML | C# keep hidden columns when saving workbook as HTML | HtmlSaveOptions HiddenColDisplayType example | How to hide columns C and D in Aspose.Cells | Export Excel to HTML with hidden columns retained
// Developer Intent: Hide columns C and D in a worksheet and generate an HTML file that retains those columns as hidden.
// Use Cases: Create a web‑ready spreadsheet view that omits sensitive columns from display while preserving them for internal logic. | Generate printable HTML reports where certain columns are concealed for end‑users but remain in the source file. | Provide a preview of a workbook on a portal, keeping hidden columns available for later unhide or data processing.
// AI Prompts: Generate C# code with Aspose.Cells to hide columns B‑E and export to HTML while preserving hidden state. | Explain the effect of HtmlHiddenColDisplayType.Hidden on the resulting HTML and how to switch to showing hidden columns. | Give step‑by‑step instructions for hiding non‑adjacent columns and exporting the workbook to HTML with those columns still hidden.

using System;
using Aspose.Cells;

// Loads an existing workbook, hides columns C and D using Cells.HideColumns, configures HtmlSaveOptions with HiddenColDisplayType.Hidden, and saves the file as HTML while keeping the hidden columns in the output.
class HideColumnsAndExport
{
    static void Main()
    {
        // Load an existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Hide columns C and D (zero‑based indexes 2 and 3)
        // HideColumns(startColumn, totalColumns) hides a range of columns
        cells.HideColumns(2, 2);

        // Prepare HTML save options to keep hidden columns in the output
        HtmlSaveOptions saveOptions = new HtmlSaveOptions
        {
            // "Hidden" keeps the columns in the HTML but marks them as hidden
            HiddenColDisplayType = HtmlHiddenColDisplayType.Hidden
        };

        // Export the workbook to HTML while preserving hidden columns
        workbook.Save("output.html", saveOptions);
    }
}
