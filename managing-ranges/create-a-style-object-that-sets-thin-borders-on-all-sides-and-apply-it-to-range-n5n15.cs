// Title: How to create a thin border style and apply it to cells N5:N15 using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code with Aspose.Cells that defines a Style with thin borders on all sides and applies it only to the range N5:N15. | Show how to use StyleFlag to restrict formatting to borders when applying a style to a specific column range in an Aspose.Cells workbook.
// Common Searches: Aspose.Cells C# set thin border for range N5 to N15 | Apply only border formatting to a column range with Aspose.Cells .NET | Create and apply a style with thin borders to cells in column N using Aspose.Cells | How to use StyleFlag to apply border style without affecting other cell formatting in Aspose.Cells | C# Aspose.Cells create range and set border style for multiple rows
// Tags: Aspose.Cells thin border style C# | apply border style to range N5:N15 Aspose.Cells | StyleFlag border-only formatting Aspose.Cells | create range N5 N15 Aspose.Cells | C# workbook border formatting Aspose.Cells

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// // This program creates a new workbook, defines a Style with thin borders on all sides, and applies that border style exclusively to the cells in column N rows 5 through 15 before saving the file as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Create a style with thin borders on all sides
            Style style = workbook.CreateStyle();
            style.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
            style.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
            style.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
            style.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;

            // Apply the style to the range N5:N15
            AsposeRange range = sheet.Cells.CreateRange("N5", "N15");
            StyleFlag flag = new StyleFlag
            {
                Borders = true // Apply only border formatting
            };
            range.ApplyStyle(style, flag);

            // Save the workbook
            workbook.Save("output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
