// Title: Create an XLSX workbook, merge cells P2:P5, enable wrap text, and apply an AutoFilter using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that merges the range P2:P5, applies a wrap‑text style to the merged cell, sets an AutoFilter covering the needed columns, and saves the result as an XLSX file. | Show how to define a style that wraps text, assign it to a merged cell, configure the AutoFilter range, and export the workbook with Aspose.Cells in a .NET project.
// Common Searches: Aspose.Cells .NET merge cells P2 to P5 and wrap text | How to add an AutoFilter that includes merged cells using C# Aspose.Cells | Saving a workbook with merged range and wrap‑text style as XLSX in Aspose.Cells | C# example for setting wrap text on a merged cell and applying AutoFilter with Aspose.Cells
// Tags: merge cells range P2:P5 Aspose.Cells | wrap text style merged cell Aspose.Cells | set autofilter range A1:P5 Aspose.Cells | save workbook as XLSX Aspose.Cells

using Aspose.Cells;
using System;

// The program creates a new workbook, merges cells P2:P5, applies a wrap‑text style to the merged cell, defines an AutoFilter covering A1:P5, and saves the file as output.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Merge cells P2:P5 (zero‑based indices: row 1, column 15)
        sheet.Cells.Merge(1, 15, 4, 1);

        // Create a style with wrap text enabled
        Style wrapStyle = workbook.CreateStyle();
        wrapStyle.IsTextWrapped = true;

        // Apply the wrap style to the merged cell (top‑left cell of the range)
        sheet.Cells["P2"].SetStyle(wrapStyle);

        // Apply an AutoFilter that includes the merged cells.
        // Here we use a fixed range that certainly covers P2:P5.
        sheet.AutoFilter.Range = "A1:P5";

        // Save the workbook as an XLSX file
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}
