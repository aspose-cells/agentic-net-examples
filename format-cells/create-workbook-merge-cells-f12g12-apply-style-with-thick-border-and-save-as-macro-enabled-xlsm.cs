// Title: Merge cells F12:G12, apply a thick border style, and save as a macro‑enabled XLSM workbook using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# merge specific cells and set thick borders | How to create a macro‑enabled XLSM file with merged cells using Aspose.Cells | Apply style to a merged range F12:G12 in Aspose.Cells .NET | Saving a workbook as XLSM after formatting merged cells in C# | Set border formatting for merged cells with Aspose.Cells
// Tags: cell merging with thick borders Aspose.Cells | macro enabled XLSM export Aspose.Cells | range styling merged cells Aspose.Cells | border formatting merged range Aspose.Cells | create workbook merged cells Aspose.Cells

using System;
using Aspose.Cells;

// The example creates a new workbook, merges cells F12:G12 on the first worksheet, defines a style with thick borders on all sides, applies this style to the merged range, and saves the workbook as a macro‑enabled XLSM file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Merge cells F12:G12 (zero‑based indices: row 11, column 5 to column 6)
            sheet.Cells.Merge(11, 5, 1, 2);

            // Create a style with thick borders on all sides
            Style thickBorderStyle = workbook.CreateStyle();
            thickBorderStyle.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thick;
            thickBorderStyle.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thick;
            thickBorderStyle.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thick;
            thickBorderStyle.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thick;

            // Apply the style to the merged range
            Aspose.Cells.Range mergedRange = sheet.Cells.CreateRange("F12:G12");
            StyleFlag flag = new StyleFlag { All = true };
            mergedRange.ApplyStyle(thickBorderStyle, flag);

            // Save the workbook as a macro‑enabled XLSM file
            string outputPath = "MergedCells.xlsm";
            workbook.Save(outputPath, SaveFormat.Xlsm);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
