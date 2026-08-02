// Title: Validate TextBox Character Spacing (‑10 to 10 pt) in Aspose.Cells for .NET
// Description: C# example that creates a workbook, adds a TextBox shape, checks a character‑spacing value against the allowed -10 to 10 point range, throws an ArgumentOutOfRangeException for invalid values, and saves the file when the spacing is valid.
// Keywords: Aspose.Cells | character spacing | TextBox shape | .NET | validation | ArgumentOutOfRangeException | font spacing range | C# Excel example | shape formatting | spacing limits
// Common Searches: Aspose.Cells validate character spacing | TextBox spacing range Aspose.Cells .NET | how to limit font spacing to -10 to 10 points | catch ArgumentOutOfRangeException for shape spacing | C# example checking text box character spacing
// Developer Intent: Confirm that a character‑spacing value for a TextBox shape falls within the supported -10 to 10 point interval before applying it.
// Use Cases: Pre‑validate spacing to prevent runtime errors when formatting shapes. | Provide user‑friendly error messages for out‑of‑range spacing inputs. | Integrate spacing checks into reusable utility methods for any Aspose.Cells shape.
// AI Prompts: Write C# code that verifies a character spacing value is between -10 and 10 points before assigning it to a TextBox in Aspose.Cells. | Show how to handle ArgumentOutOfRangeException when an invalid spacing value is supplied to a shape. | Create a generic method that validates character spacing for all shape types in Aspose.Cells and returns a boolean result.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // C# example that creates a workbook, adds a TextBox shape, checks a character‑spacing value against the allowed -10 to 10 point range, throws an ArgumentOutOfRangeException for invalid values, and saves the file when the spacing is valid.
    public class ValidateCharacterSpacing
    {
        // Allowed spacing range in points
        private const double MinSpacing = -10.0;
        private const double MaxSpacing = 10.0;

        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a text box shape to the worksheet
                TextBox textBox = worksheet.Shapes.AddTextBox(0, 0, 2, 0, 200, 100);
                textBox.Text = "Sample Text for Spacing Validation";

                // Example spacing value to validate (intentionally out of range)
                double spacingToSet = 12.5;

                // Validate the spacing value
                if (spacingToSet < MinSpacing || spacingToSet > MaxSpacing)
                {
                    throw new ArgumentOutOfRangeException(
                        nameof(spacingToSet),
                        spacingToSet,
                        $"Spacing must be between {MinSpacing} and {MaxSpacing} points.");
                }

                // NOTE: Aspose.Cells Font does not expose a CharacterSpacing property.
                // If needed, implement custom handling here. For now, we simply acknowledge the valid value.
                // Example: textBox.Font.Size = (float)(textBox.Font.Size + spacingToSet * 0.1);

                // Save the workbook
                string outputPath = "ValidatedSpacing.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (ArgumentOutOfRangeException ex)
            {
                Console.WriteLine($"Invalid spacing value: {ex.Message}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            ValidateCharacterSpacing.Run();
        }
    }
}
