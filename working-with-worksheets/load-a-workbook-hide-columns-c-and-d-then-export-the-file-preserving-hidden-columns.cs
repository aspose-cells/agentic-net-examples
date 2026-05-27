using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load an existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Hide columns C and D (zero‑based indexes 2 and 3)
        // HideColumns(startColumn, totalColumns) hides a range of columns
        workbook.Worksheets[0].Cells.HideColumns(2, 2);

        // Prepare HTML save options to keep hidden columns in the output
        HtmlSaveOptions saveOptions = new HtmlSaveOptions
        {
            // "Hidden" keeps the column in the HTML but marks it as hidden (preserves layout)
            HiddenColDisplayType = HtmlHiddenColDisplayType.Hidden
        };

        // Export the workbook to HTML while preserving hidden columns
        workbook.Save("output.html", saveOptions);
    }
}