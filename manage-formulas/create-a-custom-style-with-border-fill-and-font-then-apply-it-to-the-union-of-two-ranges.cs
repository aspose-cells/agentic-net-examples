// Title: Create a custom cell style with borders, fill, and font and apply it to a union of two ranges using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# apply a custom style to multiple non‑contiguous cells | How to create and apply a style with borders, background color, and font to a union range in Aspose.Cells | Using StyleFlag to apply formatting to a range created with CreateRange in Aspose.Cells .NET | Apply yellow fill and blue bold Arial font to A1:C3 and E5:G7 with Aspose.Cells | Define a custom cell style and apply it to a union of ranges in a workbook using Aspose.Cells for .NET
// Tags: custom cell style with borders fill font Aspose.Cells | apply style to union range Aspose.Cells | StyleFlag All property usage | CreateRange non‑contiguous range Aspose.Cells | save workbook as XLSX Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;

// The example creates a new workbook, defines a custom style with thin borders on all sides, a solid yellow background, and a bold blue Arial 12‑point font, builds a union range covering A1:C3 and E5:G7, applies the style to the entire union using a StyleFlag with All=true, and saves the workbook as StyledUnionRanges.xlsx.
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

            // -----------------------------
            // Create a custom style
            // -----------------------------
            Style customStyle = workbook.CreateStyle();

            // Set borders (thin on all sides)
            customStyle.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
            customStyle.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
            customStyle.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
            customStyle.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;

            // Set fill (solid yellow background)
            customStyle.ForegroundColor = Color.Yellow;
            customStyle.Pattern = BackgroundType.Solid;

            // Set font (Arial, 12pt, blue, bold)
            customStyle.Font.Name = "Arial";
            customStyle.Font.Size = 12;
            customStyle.Font.Color = Color.Blue;
            customStyle.Font.IsBold = true;

            // -----------------------------
            // Define two separate ranges
            // -----------------------------
            // First range: A1:C3
            // Second range: E5:G7
            // Create a union range that includes both areas
            Aspose.Cells.Range unionRange = sheet.Cells.CreateRange("A1:C3,E5:G7");

            // -----------------------------
            // Apply the custom style to the union range
            // -----------------------------
            // Use StyleFlag to apply all style elements
            StyleFlag flag = new StyleFlag() { All = true };
            unionRange.ApplyStyle(customStyle, flag);

            // -----------------------------
            // Save the workbook
            // -----------------------------
            workbook.Save("StyledUnionRanges.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
