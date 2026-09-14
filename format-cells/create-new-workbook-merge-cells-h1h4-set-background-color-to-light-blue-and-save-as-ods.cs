// Title: How to merge cells H1:H4, apply a light‑blue fill, and export to ODS with Aspose.Cells for .NET
// AI Prompts: Generate C# code using Aspose.Cells that merges the range H1:H4, sets a solid light‑blue background, and saves the workbook as an ODS file. | Show a step‑by‑step example of creating a style, applying it to a merged cell range, and exporting the result to ODS in a .NET application.
// Common Searches: Aspose.Cells C# merge H1:H4 and set background color before saving as ODS | Create ODS file with merged cells and light blue fill using Aspose.Cells .NET | How to apply solid fill to a merged range in Aspose.Cells for .NET | Saving a workbook with styled merged cells to ODS format in C#
// Tags: merge cells range Aspose.Cells C# | apply solid fill style Aspose.Cells | export workbook to ODS Aspose.Cells | style merged range light blue Aspose | create ODS document with colored merged cells .NET

using System;
using System.Drawing;
using Aspose.Cells;

// Creates a new workbook, merges cells H1:H4, applies a solid light‑blue background style to the merged range, and saves the file as an ODS document using Aspose.Cells for .NET.
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

            // Merge cells H1:H4 (row 0-3, column 7)
            sheet.Cells.Merge(0, 7, 3, 7);

            // Create a style with light blue background
            Style style = workbook.CreateStyle();
            style.ForegroundColor = Color.LightBlue;
            style.Pattern = BackgroundType.Solid;

            // Apply the style to the merged range
            Aspose.Cells.Range range = sheet.Cells.CreateRange("H1:H4");
            StyleFlag flag = new StyleFlag
            {
                CellShading = true // Apply background color
            };
            range.ApplyStyle(style, flag);

            // Save the workbook as ODS (use the non‑obsolete enum value)
            workbook.Save("MergedLightBlue.ods", SaveFormat.Ods);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
