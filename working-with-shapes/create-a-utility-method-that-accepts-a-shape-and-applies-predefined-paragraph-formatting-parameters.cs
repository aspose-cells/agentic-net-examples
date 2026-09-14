// Title: Apply paragraph formatting (alignment, font, placeholder) to an Aspose.Cells Shape in C#
// AI Prompts: Create a C# helper method that takes an Aspose.Cells Shape, validates the argument, inserts "Sample Text" when the shape's Text is empty, centers the paragraph, and sets the font to Calibri 11 pt black using the Shape.Font API. | Write code that adds a rectangle shape to a worksheet, then calls a utility to apply horizontal alignment and font styling to the shape's text, while handling missing template files and ensuring the output directory exists.
// Common Searches: c# aspocells set shape text alignment to center | aspocells add placeholder text to shape if empty | how to change font of shape text in aspocells workbook | utility method for formatting shape paragraph in aspocells .net | apply paragraph formatting to rectangle shape using aspocells drawing api
// Tags: shape paragraph formatting aspocells c# | apply text alignment to aspocells shape | set shape font properties aspocells | add default text to empty shape aspocells | c# aspocells shape helper method

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsUtilities
{
    // The ShapeHelper class defines ApplyParagraphFormatting, which validates a Shape, adds placeholder text when needed, centers the paragraph, and applies Calibri 11‑point black font settings. The sample program loads or creates a workbook, inserts a rectangle shape, invokes the helper, and saves the result to Result.xlsx.
    public static class ShapeHelper
    {
        /// <param name="shape">The shape whose text will be formatted.</param>
        public static void ApplyParagraphFormatting(Shape shape)
        {
            if (shape == null)
                throw new ArgumentNullException(nameof(shape));

            // Ensure the shape contains some text; if empty, add a placeholder.
            if (string.IsNullOrEmpty(shape.Text))
                shape.Text = "Sample Text";

            // Horizontal alignment
            shape.TextHorizontalAlignment = TextAlignmentType.Center;

            // Font formatting (use Font property for compatibility with older Aspose.Cells versions)
            Font font = shape.Font;
            font.Name = "Calibri";
            font.Size = 11;
            font.IsBold = false;
            font.IsItalic = false;
            font.Color = Color.Black;
        }
    }

    class Program
    {
        static void Main()
        {
            try
            {
                // Path to an optional template workbook.
                string templatePath = "Template.xlsx";

                Workbook workbook;
                if (File.Exists(templatePath))
                {
                    // Load existing workbook.
                    workbook = new Workbook(templatePath);
                }
                else
                {
                    // Create a new workbook if template is missing.
                    workbook = new Workbook();
                }

                // Access the first worksheet.
                Worksheet sheet = workbook.Worksheets[0];

                // Add a rectangle shape to the worksheet.
                // Parameters order: type, upperLeftRow, upperLeftColumn, upperLeftRowOffset,
                // upperLeftColumnOffset, height, width.
                Shape shape = sheet.Shapes.AddShape(
                    MsoDrawingType.Rectangle,
                    2,    // upperLeftRow
                    1,    // upperLeftColumn
                    0,    // upperLeftRowOffset (pixels)
                    0,    // upperLeftColumnOffset (pixels)
                    100,  // height (pixels)
                    200   // width (pixels)
                );
                shape.Name = "DemoShape";

                // Apply formatting to the shape's text.
                ShapeHelper.ApplyParagraphFormatting(shape);

                // Ensure output directory exists.
                string outputPath = "Result.xlsx";
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the result.
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
