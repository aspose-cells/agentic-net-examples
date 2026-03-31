using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the source XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Apply a single strikethrough to cell A1
        Cell cell = workbook.Worksheets[0].Cells["A1"];
        cell.PutValue("Strikethrough Text");
        Style style = cell.GetStyle();
        style.Font.IsStrikeout = true; // enable single strikeout
        cell.SetStyle(style);

        // Configure HTML save options
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        // Use the Cross type for cross‑cell strings (faster rendering)
        htmlOptions.HtmlCrossStringType = HtmlCrossType.Cross;

        // Save the worksheet as HTML with the configured options
        workbook.Save("output.html", htmlOptions);
    }
}