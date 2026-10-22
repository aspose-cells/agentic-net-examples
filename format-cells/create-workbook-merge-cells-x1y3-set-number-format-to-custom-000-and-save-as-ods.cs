// Title: Merge cells X1:Y3, apply a custom Euro number format, and save as ODS with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells merge X1:Y3 and set Euro currency format | How to apply custom '#,##0.00 €' number format to merged cells before exporting to ODS using Aspose.Cells | Saving a merged cell range with custom Euro formatting to ODS in Aspose.Cells .NET | Aspose.Cells example for merging cells and applying Euro number format prior to ODS export | Create an ODS file with merged cells and a custom Euro number format using Aspose.Cells C#
// Tags: merge cell range X1:Y3 Aspose.Cells | custom Euro currency format Aspose.Cells | export workbook to ODS format Aspose.Cells | apply number format to merged cells Aspose.Cells | style merged range Aspose.Cells C#

using Aspose.Cells;
using System;

// The example creates a new workbook, merges the range X1:Y3, applies the custom number format "#,##0.00 €" to the merged cells, and saves the workbook as an ODS file using Aspose.Cells for .NET.
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

            // Merge cells from X1 to Y3 (rows 0-2, columns 23-24)
            sheet.Cells.Merge(0, 23, 2, 24);

            // Create a style with the custom number format "#,##0.00 €"
            Style customStyle = workbook.CreateStyle();
            customStyle.Custom = "#,##0.00 €";

            // Apply the style to the merged range
            Aspose.Cells.Range mergedRange = sheet.Cells.CreateRange("X1:Y3");
            StyleFlag flag = new StyleFlag();
            flag.NumberFormat = true;
            mergedRange.ApplyStyle(customStyle, flag);

            // Save the workbook as ODS (use the updated enum value)
            workbook.Save("output.ods", SaveFormat.Ods);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
