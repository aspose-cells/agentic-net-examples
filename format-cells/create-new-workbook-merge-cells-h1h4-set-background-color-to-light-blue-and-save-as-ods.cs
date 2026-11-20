// Title: Merge cells H1:H4, apply a light‑blue fill, and save the workbook as ODS using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# merge a column range and set solid fill | Create an ODS file with merged cells using .NET | Set background color for a merged range in Aspose.Cells | Save a styled workbook as ODS format in C# | Apply light blue shading to merged cells with Aspose.Cells API
// Tags: merge cells H1:H4 Aspose.Cells | solid light blue cell style Aspose.Cells | ODS output with styled merged range | Aspose.Cells workbook formatting C# | range styling and ODS export Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

// Creates a new workbook, merges cells H1:H4, applies a solid light‑blue fill to the merged range, and saves the file as an ODS document using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Merge cells H1:H4 (column index 7, rows 0-3)
            sheet.Cells.Merge(0, 7, 4, 1);

            // Create a style with light blue background
            Style style = workbook.CreateStyle();
            style.ForegroundColor = Color.LightBlue;
            style.Pattern = BackgroundType.Solid;

            // Apply the style to the merged range
            StyleFlag flag = new StyleFlag { CellShading = true };
            Aspose.Cells.Range mergedRange = sheet.Cells.CreateRange(0, 7, 4, 1);
            mergedRange.ApplyStyle(style, flag);

            // Define output file path
            string outputPath = "output.ods";

            // Save the workbook as ODS (use the non‑obsolete enum value)
            workbook.Save(outputPath, SaveFormat.Ods);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
