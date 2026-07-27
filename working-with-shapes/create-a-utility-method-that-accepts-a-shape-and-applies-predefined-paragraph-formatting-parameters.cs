// Title: C# Utility to Apply Paragraph Formatting and Auto‑Fit Shapes with Aspose.Cells
// Description: Reusable C# method that loops through all TextParagraph objects in a Shape’s TextBody, sets line spacing, alignment, margins, first‑line indent, and spacing before/after, then calls FitToTextSize to resize the shape. Includes a full example that creates a workbook, adds a text box, applies the formatting, and saves the file.
// Keywords: Aspose.Cells shape formatting | C# paragraph formatting Aspose.Cells | FitToTextSize Aspose.Cells | text box line spacing .NET | shape margins Aspose.Cells | Excel shape paragraph style | Aspose.Cells TextParagraph API | C# Excel shape utilities
// Common Searches: Aspose.Cells set paragraph line spacing in a shape | center text inside a text box using Aspose.Cells .NET | auto resize shape to fit text Aspose.Cells | apply left and right margins to shape paragraphs Aspose.Cells | custom paragraph formatting for Excel shapes C#
// Developer Intent: The developer wants a simple, reusable way to apply consistent paragraph styling to any shape’s text body and have the shape automatically adjust its size to the formatted content.
// Use Cases: Standardize multi‑paragraph text boxes in automatically generated Excel reports. | Apply uniform paragraph styles to shapes used as comments, annotations, or callouts across worksheets. | Maintain layout integrity by auto‑resizing shapes after paragraph formatting changes.
// AI Prompts: Generate a C# method that sets line spacing, alignment, margins, first‑line indent, and spacing before/after for all paragraphs in an Aspose.Cells Shape and then calls FitToTextSize. | Show how to refactor the utility to accept custom formatting parameters (e.g., spacing, alignment, margins) instead of hard‑coded values. | Provide sample code that creates a workbook, adds a text box shape, invokes the formatting utility, and saves the workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

// Reusable C# method that loops through all TextParagraph objects in a Shape’s TextBody, sets line spacing, alignment, margins, first‑line indent, and spacing before/after, then calls FitToTextSize to resize the shape. Includes a full example that creates a workbook, adds a text box, applies the formatting, and saves the file.
public static class ShapeFormattingUtility
{
    // Applies predefined paragraph formatting to the given shape.
    public static void ApplyParagraphFormatting(Shape shape)
    {
        if (shape == null) return;

        // Iterate through all paragraphs in the shape's text body.
        foreach (TextParagraph paragraph in shape.TextBody.TextParagraphs)
        {
            // Set line spacing to 20 points.
            paragraph.LineSpaceSizeType = LineSpaceSizeType.Points;
            paragraph.LineSpace = 20;

            // Center align the paragraph.
            paragraph.AlignmentType = TextAlignmentType.Center;

            // Set left and right margins.
            paragraph.LeftMargin = 5;
            paragraph.RightMargin = 5;

            // Set first line indent.
            paragraph.FirstLineIndent = 10;

            // Set space before and after the paragraph (default unit is points).
            paragraph.SpaceBefore = 5;
            paragraph.SpaceAfter = 5;
        }

        // Adjust the shape size to fit the new formatting.
        shape.FitToTextSize();
    }
}

// Example usage of the utility method.
public class Example
{
    public static void Run()
    {
        try
        {
            // Create a new workbook.
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a text box shape to the worksheet.
            int textBoxIndex = worksheet.TextBoxes.Add(10, 10, 200, 100);
            TextBox textBox = worksheet.TextBoxes[textBoxIndex];
            textBox.Text = "First paragraph\nSecond paragraph";

            // Apply predefined paragraph formatting to the text box.
            ShapeFormattingUtility.ApplyParagraphFormatting(textBox);

            // Save the workbook.
            workbook.Save("FormattedShape.xlsx");
            Console.WriteLine("Workbook saved successfully as FormattedShape.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

// Entry point for the application.
public class Program
{
    public static void Main(string[] args)
    {
        Example.Run();
    }
}
