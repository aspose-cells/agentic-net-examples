// Title: Apply a custom style with font, fill, and borders to a UnionRange of two non‑contiguous ranges using Aspose.Cells for .NET
// AI Prompts: Generate C# code that creates a Style with Arial 12 pt blue bold font, light‑yellow fill, and thin black borders, then applies it to a UnionRange formed from ranges A1:B2 and D4:E5 in an Aspose.Cells workbook. | Show how to combine two separate cell ranges into a UnionRange and format the combined area with a predefined Style using Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# how to style non‑contiguous cells with a single custom style | C# create UnionRange and apply formatting in Aspose.Cells workbook | apply font fill and border style to multiple ranges Aspose.Cells .NET example
// Tags: custom style creation Aspose.Cells .NET | unionrange formatting Aspose.Cells | apply style to noncontiguous ranges C# | cell border and fill settings Aspose.Cells | styleflag all true Aspose.Cells example

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// Demonstrates creating a reusable Style (Arial 12 pt blue bold font, light‑yellow background, thin black borders), forming a UnionRange from A1:B2 and D4:E5, applying the style with a StyleFlag, and saving the workbook as an .xlsx file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Sample data (optional, just to have visible cells)
            sheet.Cells["A1"].PutValue("Item1");
            sheet.Cells["B2"].PutValue("Item2");
            sheet.Cells["D4"].PutValue("Item3");
            sheet.Cells["E5"].PutValue("Item4");

            // -----------------------------
            // Create a custom style
            // -----------------------------
            Style customStyle = workbook.CreateStyle();

            // Font settings
            customStyle.Font.Name = "Arial";
            customStyle.Font.Size = 12;
            customStyle.Font.Color = Color.Blue;
            customStyle.Font.IsBold = true;

            // Fill settings (solid light yellow background)
            customStyle.ForegroundColor = Color.LightYellow;
            customStyle.Pattern = BackgroundType.Solid;

            // Border settings (thin black border on all sides)
            foreach (BorderType bt in new[] { BorderType.TopBorder, BorderType.BottomBorder, BorderType.LeftBorder, BorderType.RightBorder })
            {
                customStyle.Borders[bt].LineStyle = CellBorderType.Thin;
                customStyle.Borders[bt].Color = Color.Black;
            }

            // -----------------------------
            // Define two ranges and union them
            // -----------------------------
            AsposeRange range1 = sheet.Cells.CreateRange("A1:B2");
            AsposeRange range2 = sheet.Cells.CreateRange("D4:E5");

            // Union of the two ranges (returns UnionRange)
            UnionRange unionRange = range1.UnionRanges(new[] { range2 });

            // Apply the custom style to the union range
            StyleFlag flag = new StyleFlag { All = true };
            unionRange.ApplyStyle(customStyle, flag);

            // -----------------------------
            // Save the workbook
            // -----------------------------
            string outputPath = "StyledRanges.xlsx";

            // Ensure the directory exists before saving
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath)) ?? Directory.GetCurrentDirectory();
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
