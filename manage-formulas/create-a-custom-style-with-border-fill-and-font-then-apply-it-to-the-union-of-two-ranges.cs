// Title: Create a custom cell style with borders, fill, and font and apply it to a union of ranges using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# apply the same custom style to cells A1:B2 and D4:E5 | How to define a union range and set font, fill, and border in Aspose.Cells | Create a reusable style with Arial blue bold font and yellow fill in Aspose.Cells .NET | Using StyleFlag to apply all formatting options to a multi‑area range in C# | Example of applying thick red borders to non‑contiguous ranges with Aspose.Cells
// Tags: style creation with borders fill font Aspose.Cells | multi‑area range styling using CreateRange C# | apply full formatting with StyleFlag All Aspose.Cells | non‑contiguous range formatting Aspose.Cells .NET | save styled workbook as XLSX Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// The example creates a workbook, defines a style with Arial 12 blue bold font, yellow solid fill, and thick red borders on all sides, builds a multi‑area range covering A1:B2 and D4:E5, applies the style to the entire union using a StyleFlag with All=true, and saves the result as StyledUnionRanges.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // -----------------------------
            // Define a custom style
            // -----------------------------
            Style customStyle = workbook.CreateStyle();

            // Font settings
            customStyle.Font.Name = "Arial";
            customStyle.Font.Size = 12;
            customStyle.Font.Color = Color.Blue;
            customStyle.Font.IsBold = true;

            // Fill settings
            customStyle.ForegroundColor = Color.Yellow;
            customStyle.Pattern = BackgroundType.Solid;

            // Border settings (all sides)
            customStyle.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thick;
            customStyle.Borders[BorderType.TopBorder].Color = Color.Red;

            customStyle.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thick;
            customStyle.Borders[BorderType.BottomBorder].Color = Color.Red;

            customStyle.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thick;
            customStyle.Borders[BorderType.LeftBorder].Color = Color.Red;

            customStyle.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thick;
            customStyle.Borders[BorderType.RightBorder].Color = Color.Red;

            // -----------------------------
            // Define the union of two ranges
            // Aspose.Cells allows creating a multi‑area range using a comma‑separated address string
            // -----------------------------
            AsposeRange unionRange = sheet.Cells.CreateRange("A1:B2,D4:E5");

            // Apply the custom style to the entire union range
            StyleFlag flag = new StyleFlag() { All = true };
            unionRange.ApplyStyle(customStyle, flag);

            // -----------------------------
            // Save the workbook
            // -----------------------------
            string outputPath = "StyledUnionRanges.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
