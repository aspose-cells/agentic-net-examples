// Title: Validate and truncate custom textbox label text to Excel’s 32,767‑character limit using Aspose.Cells for .NET
// Common Searches: c# Aspose.Cells check if textbox label exceeds Excel’s maximum characters before assignment | how to enforce Excel shape text length limit when adding a TextBox with Aspose.Cells | truncate string to 32767 characters for Aspose.Cells TextBox in .NET | catch ArgumentException for overlong shape text using Aspose.Cells | prevent Excel error by validating custom label length in Aspose.Cells workbook
// Tags: validate textbox text length Aspose.Cells | Excel shape text character limit .NET | truncate long label Aspose.Cells TextBox | AddTextBox text length validation | handle ArgumentException shape text Aspose.Cells | custom label size check before workbook save

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace CustomLabelValidationDemo
{
    // The example defines Excel's 32,667‑character text limit, provides a ValidateLabelLength method that throws on null or over‑length strings, demonstrates assigning a short label to a TextBox, catches the exception for an excessively long label, truncates it to the allowed size, assigns the truncated text to another TextBox, and saves the workbook as an XLSX file.
    class Program
    {
        // Excel's maximum allowed characters for a cell or shape text.
        private const int ExcelMaxTextLength = 32767;

        /// <param name="text">The label text to validate.</param>
        private static void ValidateLabelLength(string text)
        {
            if (text == null)
                throw new ArgumentNullException(nameof(text));

            if (text.Length > ExcelMaxTextLength)
                throw new ArgumentException(
                    $"Label text length ({text.Length}) exceeds Excel's maximum allowed characters ({ExcelMaxTextLength}).",
                    nameof(text));
        }

        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook.
                Workbook workbook = new Workbook();

                // Access the first worksheet.
                Worksheet worksheet = workbook.Worksheets[0];
                worksheet.Name = "Demo";

                // Example 1: Assign a valid label (short text).
                string shortLabel = "This is a short custom label.";
                try
                {
                    ValidateLabelLength(shortLabel);
                    // Add a textbox shape to the worksheet.
                    TextBox tb = worksheet.Shapes.AddTextBox(1, 1, 0, 0, 200, 50);
                    tb.Text = shortLabel;
                    Console.WriteLine("Short label assigned successfully.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to assign short label: {ex.Message}");
                }

                // Example 2: Attempt to assign an overly long label.
                // Build a string longer than Excel's limit.
                string longLabel = new string('X', ExcelMaxTextLength + 10);
                try
                {
                    ValidateLabelLength(longLabel);
                    // This line will not be reached because the validation throws.
                    TextBox tb = worksheet.Shapes.AddTextBox(5, 1, 0, 0, 300, 100);
                    tb.Text = longLabel;
                    Console.WriteLine("Long label assigned successfully (unexpected).");
                }
                catch (ArgumentException argEx)
                {
                    Console.WriteLine($"Validation caught: {argEx.Message}");
                    // Optionally truncate and assign the allowed portion.
                    string truncated = longLabel.Substring(0, ExcelMaxTextLength);
                    TextBox tb = worksheet.Shapes.AddTextBox(5, 1, 0, 0, 300, 100);
                    tb.Text = truncated;
                    Console.WriteLine($"Truncated label assigned with length {truncated.Length}.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Unexpected error: {ex.Message}");
                }

                // Save the workbook to disk.
                string outputPath = "CustomLabelValidation.xlsx";
                try
                {
                    workbook.Save(outputPath, SaveFormat.Xlsx);
                    Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to save workbook: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Fatal error: {ex.Message}");
            }
        }
    }
}
