using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    public class ValidateCharacterSpacing
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a text box shape to the worksheet
                TextBox textBox = worksheet.Shapes.AddTextBox(0, 0, 2, 0, 200, 100);
                textBox.Text = "Sample text for spacing validation";

                // Desired character spacing (points). Aspose.Cells Font does not expose a CharacterSpacing property,
                // so we only validate the value here without applying it to the font.
                double desiredSpacing = 12.5;

                // Validate that the spacing is within the allowed range [-10, 10] points
                if (desiredSpacing < -10.0 || desiredSpacing > 10.0)
                {
                    throw new ArgumentOutOfRangeException(
                        nameof(desiredSpacing),
                        $"Spacing value {desiredSpacing} is outside the allowed range of -10 to 10 points.");
                }

                // Save the workbook
                string outputPath = "ValidatedSpacing.xlsx";

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (ArgumentOutOfRangeException ex)
            {
                Console.WriteLine($"Validation error: {ex.Message}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }

        // Entry point for the application
        public static void Main(string[] args)
        {
            Run();
        }
    }
}