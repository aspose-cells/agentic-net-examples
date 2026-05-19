using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

namespace AsposeCellsDemo
{
    public static class ShapeFormattingUtility
    {
        /// <summary>
        /// Applies predefined paragraph formatting to all paragraphs of the given shape.
        /// </summary>
        /// <param name="shape">The shape whose text paragraphs will be formatted.</param>
        public static void ApplyParagraphFormatting(Shape shape)
        {
            try
            {
                // Ensure the shape contains a text body.
                if (shape?.TextBody?.TextParagraphs == null)
                    return;

                foreach (TextParagraph paragraph in shape.TextBody.TextParagraphs)
                {
                    // Center alignment.
                    paragraph.AlignmentType = TextAlignmentType.Center;

                    // Line spacing (15 points).
                    paragraph.LineSpaceSizeType = LineSpaceSizeType.Points;
                    paragraph.LineSpace = 15;

                    // Left and right margins (5 points each).
                    paragraph.LeftMargin = 5;
                    paragraph.RightMargin = 5;

                    // Space before and after (8 points each). 
                    // The default size type is Points, so explicit enum assignment is unnecessary.
                    paragraph.SpaceBefore = 8;
                    paragraph.SpaceAfter = 8;

                    // First line indent (10 points).
                    paragraph.FirstLineIndent = 10;
                }

                // Recalculate the shape's text layout after formatting.
                shape.CalculateTextSize();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error while applying paragraph formatting: {ex.Message}");
            }
        }
    }

    internal class Program
    {
        private static void Main()
        {
            const string inputPath = "Template.xlsx";
            const string outputPath = "Formatted.xlsx";

            try
            {
                // Verify that the input file exists.
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook.
                var workbook = new Workbook(inputPath);
                var worksheet = workbook.Worksheets[0];

                // Find the first shape that contains a text body.
                Shape targetShape = null;
                foreach (Shape shape in worksheet.Shapes)
                {
                    if (shape.TextBody != null)
                    {
                        targetShape = shape;
                        break;
                    }
                }

                if (targetShape == null)
                {
                    Console.WriteLine("No shape with a text body was found in the worksheet.");
                    return;
                }

                // Apply paragraph formatting to the shape.
                ShapeFormattingUtility.ApplyParagraphFormatting(targetShape);

                // Save the modified workbook.
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Runtime error: {ex.Message}");
            }
        }
    }
}