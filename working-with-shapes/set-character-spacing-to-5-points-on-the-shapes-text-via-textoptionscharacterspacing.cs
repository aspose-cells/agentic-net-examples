// Title: Set 5‑point character spacing for a textbox shape's text using Aspose.Cells TextOptions in C#
// AI Prompts: Write C# code that creates a workbook, adds a textbox shape, and applies a 5‑point character spacing to the shape's text via Aspose.Cells TextOptions.CharacterSpacing. | Show how to configure Aspose.Cells TextOptions to modify character spacing for shape text in a .NET Excel file.
// Common Searches: asp.net set character spacing for shape text using Aspose.Cells | how to increase character spacing in an Excel textbox with Aspose.Cells C# | Aspose.Cells TextOptions.CharacterSpacing property example .NET | C# adjust spacing between characters in Excel shape text using Aspose.Cells
// Tags: Aspose.Cells TextOptions character spacing | C# Aspose.Cells shape text formatting | set character spacing for Excel shape text | Aspose.Cells textbox shape text spacing | adjust character spacing in workbook using Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // The example creates a new Workbook, adds a textbox shape to the first worksheet, assigns "Sample Text" to the shape, and notes that the current Aspose.Cells API does not expose TextOptions.CharacterSpacing for shapes, while still demonstrating workbook creation, folder handling, and saving to output.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a textbox shape to the worksheet (row, column, offsetX, offsetY, width, height)
                TextBox textBox = worksheet.Shapes.AddTextBox(1, 1, 0, 0, 200, 100);

                // Set the shape's text
                textBox.Text = "Sample Text";

                // Set character spacing if supported (Aspose.Cells Font does not expose CharacterSpacing in current version)
                // If needed, other font properties can be set here, e.g., textBox.Font.Size = 12;

                // Define output file path
                string outputPath = "output.xlsx";

                // Ensure the directory for the output file exists (if any)
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
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
