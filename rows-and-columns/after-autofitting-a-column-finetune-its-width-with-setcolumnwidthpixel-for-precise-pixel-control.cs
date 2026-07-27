using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook wb = new Workbook();
        Worksheet sheet = wb.Worksheets[0];

        // Sample data to demonstrate auto‑fit
        sheet.Cells[0, 0].Value = "This is a relatively long piece of text that will trigger auto‑fit.";
        sheet.Cells[1, 0].Value = "Short";

        // Auto‑fit column 0 (imprecise)
        sheet.AutoFitColumn(0);

        // Fine‑tune the column width to an exact pixel value (e.g., 120 pixels)
        sheet.Cells.SetColumnWidthPixel(0, 120);

        // Save the workbook
        wb.Save("AutoFitFineTuned.xlsx");
    }
}

// Author note: This example shows how to auto‑fit a column and then adjust its width precisely using SetColumnWidthPixel.