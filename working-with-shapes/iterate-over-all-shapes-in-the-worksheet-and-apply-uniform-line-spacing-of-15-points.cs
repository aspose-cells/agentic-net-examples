// Title: Set 1.5‑point line spacing for every rich‑text shape in an Aspose.Cells worksheet (C#)
// Description: Creates a workbook, adds a multiline textbox, iterates over all shapes on the first worksheet, identifies rich‑text shapes, and applies a 1.5‑point line‑spacing to each paragraph before saving as UniformLineSpacing.xlsx.
// Keywords: Aspose.Cells line spacing C# | Aspose.Cells shape formatting | Aspose.Cells TextParagraph line spacing | iterate worksheet shapes Aspose.Cells | Excel shape line spacing .NET | rich text shape Aspose.Cells | C# Aspose.Cells API example
// Common Searches: How to set line spacing for text boxes in Aspose.Cells .NET | Iterate over worksheet shapes and change paragraph formatting with Aspose.Cells | Apply 1.5 point line spacing to all rich text shapes in an Excel workbook using C# | Aspose.Cells change line spacing for shapes programmatically
// Developer Intent: Apply a uniform 1.5‑point line spacing to every paragraph of each rich‑text shape in a worksheet.
// Use Cases: Standardize paragraph spacing in auto‑generated Excel reports. | Ensure consistent text layout before converting workbooks to PDF or image formats. | Update legacy workbooks to match corporate style guidelines for line spacing.
// AI Prompts: Write C# code with Aspose.Cells that sets line spacing of all text shapes to 2 points. | Show how to detect rich‑text shapes and modify their paragraph properties (font, alignment, line spacing) using Aspose.Cells. | Provide an example that loops through shapes in multiple worksheets and applies a specified line‑spacing value.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

// Creates a workbook, adds a multiline textbox, iterates over all shapes on the first worksheet, identifies rich‑text shapes, and applies a 1.5‑point line‑spacing to each paragraph before saving as UniformLineSpacing.xlsx.
class ApplyUniformLineSpacing
{
    static void Main(string[] args)
    {
        try
        {
            Run();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    public static void Run()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add a textbox shape with multiline text
        Shape textBox = sheet.Shapes.AddTextBox(2, 0, 2, 0, 200, 100);
        textBox.Text = "Line one\nLine two\nLine three";

        // Iterate through all shapes in the worksheet
        for (int i = 0; i < sheet.Shapes.Count; i++)
        {
            Shape shape = sheet.Shapes[i];

            // Process only shapes that contain rich text
            if (shape.IsRichText && shape.TextBody != null)
            {
                TextParagraphCollection paragraphs = shape.TextBody.TextParagraphs;

                // Apply uniform line spacing of 1.5 points to each paragraph
                for (int p = 0; p < paragraphs.Count; p++)
                {
                    TextParagraph paragraph = paragraphs[p];
                    paragraph.LineSpaceSizeType = LineSpaceSizeType.Points;
                    paragraph.LineSpace = 1.5;
                }
            }
        }

        // Save the workbook
        workbook.Save("UniformLineSpacing.xlsx");
    }
}
