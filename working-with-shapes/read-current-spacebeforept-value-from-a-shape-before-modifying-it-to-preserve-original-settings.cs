// Title: Read a shape's original text and preserve its formatting before modifying it with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# read existing shape text before editing | preserve textbox formatting when updating shape text in Excel using Aspose.Cells | how to get and keep original shape properties while changing text with Aspose.Cells .NET | retrieve shape text without losing style Aspose.Cells C# example | read shape Text property and maintain formatting Aspose.Cells workbook
// Tags: Aspose.Cells read shape text property | preserve shape formatting during text update | modify Excel textbox content programmatically | C# Aspose.Cells shape property handling | workbook save after shape text modification

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an Excel workbook, checks for shapes on the first worksheet, reads the original Text of the first shape, outputs it, appends a modification string while keeping existing formatting, and saves the workbook to a new file, handling missing files and directories.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            var workbook = new Workbook(inputPath);

            // Access the first worksheet
            var worksheet = workbook.Worksheets[0];

            // Ensure there is at least one shape on the worksheet
            if (worksheet.Shapes.Count == 0)
            {
                Console.WriteLine("No shapes found on the worksheet.");
                return;
            }

            var shape = worksheet.Shapes[0];

            // Attempt to read and modify text for shapes that support it (e.g., TextBox)
            try
            {
                // Read the original text
                string originalText = shape.Text;
                Console.WriteLine($"Original Text: {originalText}");

                // Example modification: append a note to the text
                shape.Text = originalText + " (modified)";
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Shape does not support text manipulation or an error occurred: {ex.Message}");
            }

            // Ensure the output directory exists
            string? outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
