// Title: C# utility to apply paragraph formatting (alignment, font, line breaks, indent) to an Aspose.Cells textbox shape
// Common Searches: how to set center alignment and bold font for a textbox shape using Aspose.Cells in C# | Aspose.Cells C# collapse consecutive line breaks in shape text | add left indent to textbox shape text with Aspose.Cells .NET | apply custom paragraph formatting to a shape programmatically in Aspose.Cells | C# example of formatting shape text (alignment, font size, color) in an Excel workbook with Aspose.Cells
// Tags: aspocells shape text alignment and font styling | aspocells textbox collapse multiple line breaks | aspocells shape simulate left indent | c# aspocells shape paragraph formatting method | aspocells shape text vertical horizontal centering

using System;
using System.Drawing;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Provides a static C# method ApplyParagraphFormatting that receives an Aspose.Cells Shape, centers its text horizontally and vertically, sets the font to 12‑point bold black, collapses consecutive blank lines, and prefixes each line with four spaces to simulate a left indent. The example loads or creates a workbook, adds a textbox shape, assigns sample text, invokes the utility, and saves the workbook as FormattedShape.xlsx.
public static class ShapeFormattingUtility
{
    /// <param name="shape">The Aspose.Cells shape to format.</param>
    public static void ApplyParagraphFormatting(Shape shape)
    {
        // Horizontal and vertical alignment.
        shape.TextHorizontalAlignment = TextAlignmentType.Center;
        shape.TextVerticalAlignment = TextAlignmentType.Center;

        // Font formatting.
        shape.Font.Size = 12;               // 12 pt
        shape.Font.IsBold = true;           // bold
        shape.Font.IsItalic = false;        // not italic
        shape.Font.Color = Color.Black;     // black

        // Collapse multiple line breaks to a single one.
        if (!string.IsNullOrEmpty(shape.Text))
        {
            shape.Text = Regex.Replace(shape.Text, @"\r?\n{2,}", "\n");
        }

        // Simulate a left indent by prefixing spaces.
        if (!string.IsNullOrEmpty(shape.Text))
        {
            var lines = shape.Text.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);
            for (int i = 0; i < lines.Length; i++)
            {
                lines[i] = "    " + lines[i];
            }
            shape.Text = string.Join("\n", lines);
        }
    }
}

public class Program
{
    private const string InputFile = "Template.xlsx";
    private const string OutputFile = "FormattedShape.xlsx";

    public static void Main()
    {
        try
        {
            // Load an existing workbook if present; otherwise create a new one.
            Workbook workbook;
            if (File.Exists(InputFile))
            {
                workbook = new Workbook(InputFile);
            }
            else
            {
                workbook = new Workbook();
                workbook.Worksheets[0].Name = "Sheet1";
            }

            Worksheet sheet = workbook.Worksheets[0];

            // Add a textbox shape (supports text). Parameters: upperLeftRow, upperLeftColumn, rowOffset, columnOffset, height, width.
            Shape shape = sheet.Shapes.AddTextBox(2, 1, 0, 0, 100, 200);
            shape.Text = "Line1\nLine2\n\nLine4";

            // Apply the paragraph formatting.
            ShapeFormattingUtility.ApplyParagraphFormatting(shape);

            // Save the result.
            workbook.Save(OutputFile);
            Console.WriteLine($"Workbook saved to '{OutputFile}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
