// Title: Implement safe Shape.Characters range validation to prevent out‑of‑bounds errors in Aspose.Cells for .NET
// Common Searches: C# how to check startIndex and length before calling Shape.Characters in Aspose.Cells | Aspose.Cells example for handling ArgumentException when character range exceeds shape text length | validate shape text length before applying font settings with Aspose.Cells | prevent out of range errors with Shape.Characters method in .NET workbook | safe wrapper for GetCharacters on a shape in Aspose.Cells
// Tags: shape characters range validation Aspose.Cells | Aspose.Cells safe GetCharacters wrapper | C# argument exception handling for shape text range | out‑of‑bounds character index protection Aspose.Cells | shape text formatting error handling .NET

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using System.Drawing;

namespace AsposeCellsCharactersErrorHandling
{
    // The sample defines a GetCharactersSafe method that verifies a Shape's text, ensures startIndex and length are non‑negative and within the text length, and throws ArgumentException for invalid ranges before invoking shape.Characters. It demonstrates applying font formatting to a valid character range, catching out‑of‑range errors with try‑catch blocks, and saving the workbook to confirm the applied styling.
    class Program
    {
        /// <param name="shape">The shape containing the text.</param>
        /// <param name="startIndex">Zero‑based start index of the character range.</param>
        /// <param name="length">Number of characters in the range.</param>
        /// <returns>FontSetting object for the valid range.</returns>
        private static FontSetting GetCharactersSafe(Shape shape, int startIndex, int length)
        {
            if (shape == null) throw new ArgumentNullException(nameof(shape));

            // Shape.Text may be null; treat null as empty string.
            string text = shape.Text ?? string.Empty;

            if (startIndex < 0)
                throw new ArgumentException("startIndex cannot be negative.", nameof(startIndex));

            if (length < 0)
                throw new ArgumentException("length cannot be negative.", nameof(length));

            if (startIndex > text.Length)
                throw new ArgumentException($"startIndex ({startIndex}) exceeds the text length ({text.Length}).", nameof(startIndex));

            if (startIndex + length > text.Length)
                throw new ArgumentException($"The range (startIndex + length = {startIndex + length}) exceeds the text length ({text.Length}).", nameof(length));

            // The range is valid – call the Aspose API.
            return shape.Characters(startIndex, length);
        }

        static void Main(string[] args)
        {
            // Create a new workbook and get the first worksheet.
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a rectangle shape with some sample text.
            Shape shape = worksheet.Shapes.AddRectangle(1, 0, 0, 300, 100, 0);
            shape.Text = "Aspose.Cells Demo";

            Console.WriteLine($"Shape text: \"{shape.Text}\" (Length = {shape.Text.Length})");

            // Example 1: Valid character range.
            try
            {
                int validStart = 0;
                int validLength = 6; // "Aspose"
                FontSetting fontSetting = GetCharactersSafe(shape, validStart, validLength);
                fontSetting.Font.Color = Color.Red;
                fontSetting.Font.IsBold = true;
                fontSetting.Font.Size = 14;

                Console.WriteLine($"Applied formatting to characters [{validStart}, {validStart + validLength - 1}].");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error applying valid range: {ex.Message}");
            }

            // Example 2: Invalid character range (exceeds text length).
            try
            {
                int invalidStart = 5;
                int invalidLength = 20; // This goes beyond the text length.
                FontSetting fontSetting = GetCharactersSafe(shape, invalidStart, invalidLength);
                // The line above should throw; the following lines will not execute.
                fontSetting.Font.Color = Color.Blue;
                fontSetting.Font.IsItalic = true;
                fontSetting.Font.Size = 12;

                Console.WriteLine("This line will not be reached because the range is invalid.");
            }
            catch (ArgumentException argEx)
            {
                // Expected handling for out‑of‑range parameters.
                Console.WriteLine($"Argument error: {argEx.Message}");
            }
            catch (Exception ex)
            {
                // Any other unexpected errors.
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }

            // Save the workbook to verify that the valid formatting was applied.
            string outputPath = "ShapeCharactersErrorHandlingDemo.xlsx";
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
    }
}
