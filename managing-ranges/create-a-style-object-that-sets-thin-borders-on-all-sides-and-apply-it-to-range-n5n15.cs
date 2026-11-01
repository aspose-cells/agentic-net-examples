// Title: Create a thin border style and apply it to the N5:N15 range with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set thin borders on a specific column range | How to apply a custom style to cells N5 through N15 using Aspose.Cells | Using StyleFlag to apply borders to a range in Aspose.Cells for .NET | Create and apply a thin border style to a range in an Excel workbook with Aspose.Cells | Apply borders to range N5:N15 programmatically with Aspose.Cells C#
// Tags: Aspose.Cells thin border style | range.ApplyStyle method Aspose.Cells | StyleFlag All property usage | CreateStyle with borders Aspose.Cells | C# Excel border formatting Aspose.Cells

using System;
using Aspose.Cells;

// The example demonstrates how to create a Style with thin borders on all sides, configure a StyleFlag to apply all style attributes, and apply the style to the N5:N15 range of the first worksheet using Aspose.Cells for .NET, then save the workbook as StyledRange.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle create rule)
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Create a new style object
            Style style = workbook.CreateStyle();

            // Set thin borders on all four sides
            style.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
            style.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
            style.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
            style.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;

            // Prepare a StyleFlag to apply all style attributes
            StyleFlag flag = new StyleFlag { All = true };

            // Define the target range N5:N15 and apply the style
            Aspose.Cells.Range range = sheet.Cells.CreateRange("N5:N15");
            range.ApplyStyle(style, flag);

            // Save the workbook (lifecycle save rule)
            workbook.Save("StyledRange.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
