// Title: Merge cells X1:Y3, apply a custom Euro number format, and save the workbook as ODS using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# merge range X1:Y3 and set custom number format '#,##0.00 €' before saving as ODS | C# code to apply Euro currency format to merged cells and export to OpenDocument Spreadsheet with Aspose.Cells | Saving a workbook with merged cells and custom number format to .ods using Aspose.Cells for .NET | How to create a merged cell range and custom style in Aspose.Cells then save as ODS file | Apply custom Euro number format to a merged range and generate ODS with Aspose.Cells in C#
// Tags: merge cells range Aspose.Cells | custom Euro number format Aspose.Cells C# | export workbook to ODS Aspose.Cells | apply style to merged range Aspose.Cells | create custom style Aspose.Cells C#

using System;
using Aspose.Cells;

// The example creates a new workbook, merges the cells from X1 to Y3 on the first worksheet, defines a custom number format "#,##0.00 €" for Euro currency, applies this style to the merged range, and saves the workbook as an ODS file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Merge cells from X1 to Y3 (columns X=23, Y=24; rows 1-3 => 0‑based indices)
            // startRow, startColumn, totalRows, totalColumns
            sheet.Cells.Merge(0, 23, 3, 2);

            // Create a style with the custom number format
            Style customStyle = workbook.CreateStyle();
            customStyle.Custom = "#,##0.00 €";

            // Apply the style to the merged range
            // Use fully qualified Aspose.Cells.Range to avoid conflict with System.Range
            Aspose.Cells.Range mergedRange = sheet.Cells.CreateRange("X1", "Y3");
            StyleFlag flag = new StyleFlag
            {
                NumberFormat = true
            };
            mergedRange.ApplyStyle(customStyle, flag);

            // Save the workbook as ODS (use the non‑obsolete enum value)
            workbook.Save("MergedCells.ods", SaveFormat.Ods);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
