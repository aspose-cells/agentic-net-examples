// Title: Apply a Custom Font, Fill, and Border Style to a UnionRange in Aspose.Cells for .NET
// Description: Demonstrates how to create a Workbook, define two separate ranges, combine them with a UnionRange, build a Style that includes Calibri bold white text, a dark slate‑blue solid fill, and thick yellow borders, and apply that Style to the entire UnionRange before saving the file.
// Keywords: Aspose.Cells | C# UnionRange style | custom style Aspose.Cells | non‑contiguous cell formatting | Excel border fill font .NET | CreateStyle Aspose.Cells | SetStyle UnionRange
// Common Searches: Aspose.Cells apply style to multiple non‑contiguous ranges | How to set font, fill, and borders on a UnionRange in C# | Create UnionRange with A1:B2 and C3:D4 Aspose.Cells | Set thick yellow borders using Aspose.Cells | C# example for styling separate areas in Excel
// Developer Intent: Define a single Style (font, background, borders) and apply it to a UnionRange that merges two distinct cell blocks in an Excel workbook using Aspose.Cells for .NET.
// Use Cases: Give two tables on the same sheet a consistent header appearance for a unified report. | Highlight summary sections located in different worksheet areas while keeping identical visual cues. | Prepare a printable worksheet where multiple data blocks share the same styling to improve readability.
// AI Prompts: Generate C# code that creates a UnionRange covering E5:F6 and H7:I8, then applies a style with red bold font, light‑gray fill, and double blue borders using Aspose.Cells. | Show how to define a reusable Style object with custom font, background, and border settings and apply it to several UnionRanges in a workbook. | Explain how to modify the example to use a dashed green border and a different background color for each area within the same UnionRange.

using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a Workbook, define two separate ranges, combine them with a UnionRange, build a Style that includes Calibri bold white text, a dark slate‑blue solid fill, and thick yellow borders, and apply that Style to the entire UnionRange before saving the file.
    public class UnionRangeCustomStyleDemo
    {
        // Entry point for the console application
        public static void Main(string[] args)
        {
            try
            {
                Run();
                Console.WriteLine("Workbook saved successfully as UnionRangeCustomStyleDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data in the first range A1:B2
            worksheet.Cells["A1"].PutValue("Item 1");
            worksheet.Cells["B1"].PutValue(100);
            worksheet.Cells["A2"].PutValue("Item 2");
            worksheet.Cells["B2"].PutValue(200);

            // Populate sample data in the second range C3:D4
            worksheet.Cells["C3"].PutValue("Item 3");
            worksheet.Cells["D3"].PutValue(300);
            worksheet.Cells["C4"].PutValue("Item 4");
            worksheet.Cells["D4"].PutValue(400);

            // Create a UnionRange that covers A1:B2 and C3:D4
            // The address string uses a comma to separate the two areas
            UnionRange unionRange = workbook.Worksheets.CreateUnionRange("A1:B2,C3:D4", 0);

            // Create a custom style with font, fill, and borders
            Style style = workbook.CreateStyle();

            // Font settings
            style.Font.Name = "Calibri";
            style.Font.Size = 12;
            style.Font.IsBold = true;
            style.Font.Color = Color.White;

            // Fill settings
            style.Pattern = BackgroundType.Solid;
            style.ForegroundColor = Color.DarkSlateBlue;

            // Border settings (apply the same style to all four borders)
            style.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thick;
            style.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thick;
            style.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thick;
            style.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thick;

            style.Borders[BorderType.TopBorder].Color = Color.Yellow;
            style.Borders[BorderType.BottomBorder].Color = Color.Yellow;
            style.Borders[BorderType.LeftBorder].Color = Color.Yellow;
            style.Borders[BorderType.RightBorder].Color = Color.Yellow;

            // Apply the custom style to the entire union range
            unionRange.SetStyle(style);

            // Save the workbook to visualize the result
            workbook.Save("UnionRangeCustomStyleDemo.xlsx");
        }
    }
}
