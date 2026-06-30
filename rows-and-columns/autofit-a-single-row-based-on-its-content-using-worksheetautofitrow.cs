using System;
using Aspose.Cells;

// Author: Aspose.Cells .NET example – auto‑fits a single row based on its content
class Program
{
    static void Main()
    {
        // Create a new workbook (replace with the provided create rule if available)
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate some cells in the target row (row index 1 → second row)
        worksheet.Cells["A2"].Value = "This is a long piece of text that should cause the row height to increase when auto‑fitted.";
        Style style = worksheet.Cells["A2"].GetStyle();
        style.IsTextWrapped = true;               // Enable text wrapping so the height can expand
        worksheet.Cells["A2"].SetStyle(style);

        // Auto‑fit only the specified row
        worksheet.AutoFitRow(1);   // Row index is zero‑based

        // Save the workbook (replace with the provided save rule if available)
        workbook.Save("AutoFitRowResult.xlsx");
    }
}