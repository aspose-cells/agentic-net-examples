// Title: Compress Text in a Narrow TextBox Shape with Negative Character Spacing using Aspose.Cells for .NET
// Description: This C# example creates a workbook, adds a 100 × 50‑point textbox shape on the first worksheet, sets its text, applies a -2 pt character spacing via TextOptions.Spacing to tighten the characters, and saves the file.
// Keywords: Aspose.Cells | .NET | C# | negative character spacing | TextOptions.Spacing | textbox shape | compress text | Excel shape formatting | narrow shape | Excel automation
// Common Searches: Aspose.Cells set negative character spacing | compress text in textbox shape C# | TextOptions.Spacing example Aspose.Cells | how to tighten characters in Excel shape | create narrow textbox with custom spacing Aspose
// Developer Intent: Apply a -2 pt character spacing to a textbox shape so the text appears more compact.
// Use Cases: Fit long labels into small dashboard widgets. | Create tight‑spaced headings for chart annotations. | Design printable forms where space is limited.
// AI Prompts: Write C# code that adds a textbox shape with Aspose.Cells and sets TextOptions.Spacing to -3 points. | Explain the impact of TextOptions.Spacing on text rendering inside Excel shapes when using Aspose.Cells. | Provide error‑handling best practices for saving a workbook after modifying shape text spacing.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // This C# example creates a workbook, adds a 100 × 50‑point textbox shape on the first worksheet, sets its text, applies a -2 pt character spacing via TextOptions.Spacing to tighten the characters, and saves the file.
    public class ApplyNegativeCharacterSpacing
    {
        public static void Main(string[] args)
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
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a narrow text box shape (width 100 points, height 50 points)
            // Parameters: upper left row, upper left column, upper left row offset, upper left column offset, width, height
            Shape textBox = worksheet.Shapes.AddTextBox(0, 0, 0, 0, 100, 50);
            textBox.Text = "Compressed Text Example";

            // Set character spacing to -2 points (negative spacing compresses characters)
            // Directly use the TextOptions property without declaring a separate variable
            textBox.TextOptions.Spacing = -2.0;

            // Define output file path
            string outputPath = "NegativeCharacterSpacingDemo.xlsx";

            // Ensure the directory exists (if a directory is specified)
            string directory = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            // Save the workbook and handle any I/O errors
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
    }
}
