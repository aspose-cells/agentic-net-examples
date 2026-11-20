// Title: Create a new workbook, merge cells L5:M7, apply italic gray style, and save as XLSX using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# merge range L5:M7 and set italic font with gray background | How to apply a solid gray fill to merged cells in a .NET workbook with Aspose.Cells | Saving a workbook with merged cells and custom style to XLSX using Aspose.Cells | C# code to merge cells and apply StyleFlag for font and shading in Aspose.Cells
// Tags: merge cells range Aspose.Cells C# | apply italic font style Aspose.Cells | solid gray fill background Aspose.Cells | save workbook as XLSX Aspose.Cells | StyleFlag font shading merged range

using System;
using Aspose.Cells;
using System.Drawing;

// Creates a new Workbook, merges the range L5:M7 on the first worksheet, defines a custom Style with italic font and a solid gray fill, applies this style to the merged range using a StyleFlag, and saves the file as MergedCellsOutput.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Merge cells L5:M7 (range: L5 to M7)
            // L = column 12 (0-indexed 11), M = column 13 (0-indexed 12)
            // Row indices are 0-based, so row 5 => index 4, row 7 => index 6
            sheet.Cells.Merge(4, 11, 3, 2); // start row, start column, rows, columns

            // Create a custom style
            Style style = workbook.CreateStyle();
            style.Font.IsItalic = true;               // italic font
            style.ForegroundColor = Color.Gray;       // gray fill background
            style.Pattern = BackgroundType.Solid;     // solid fill

            // Apply the style to the merged range
            // Resolve ambiguity with Aspose.Cells.Range
            Aspose.Cells.Range mergedRange = sheet.Cells.CreateRange("L5:M7");
            mergedRange.ApplyStyle(style, new StyleFlag
            {
                Font = true,
                CellShading = true
            });

            // Export the workbook (lifecycle rule: save)
            workbook.Save("MergedCellsOutput.xlsx", SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
