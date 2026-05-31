using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    public class ApplyNegativeCharacterSpacing
    {
        public static void Main()
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
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a narrow text box shape (row, column, row offset, column offset, width, height)
                Shape textBox = worksheet.Shapes.AddTextBox(0, 0, 0, 0, 80, 200);
                textBox.Text = "Compressed Text Example";

                // Apply negative character spacing directly via the shape's TextOptions
                textBox.TextOptions.Spacing = -2.0; // bring characters closer together

                // Save the workbook
                string outputPath = "NegativeCharacterSpacingDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Catch any runtime errors (e.g., file I/O issues)
                Console.WriteLine($"Runtime error: {ex.Message}");
            }
        }
    }
}