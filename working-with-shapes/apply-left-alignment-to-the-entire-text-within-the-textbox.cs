// Title: Apply left horizontal alignment to a textbox shape in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Use Aspose.Cells in C# to set a textbox's horizontal alignment to left within an Excel worksheet. | Locate a shape named "TextBox 1" (or the first shape) and change its text alignment to left using the Aspose.Cells API. | Update an existing Excel file so that the text inside a specific textbox shape is left‑aligned with Aspose.Cells for .NET.
// Common Searches: C# Aspose.Cells left align text inside a textbox shape | how to set TextHorizontalAlignment for a shape in Aspose.Cells .NET | find textbox by name and change alignment in Excel using Aspose.Cells | apply left horizontal alignment to Excel shape text with Aspose.Cells example
// Tags: Aspose.Cells textbox left alignment C# | TextHorizontalAlignment left Aspose.Cells | Excel shape text alignment .NET | retrieve shape by name Aspose.Cells | modify textbox shape alignment C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // The example loads an Excel file, finds a textbox shape (by name "TextBox 1" or the first shape), sets its horizontal text alignment to left via the TextHorizontalAlignment property, and saves the updated workbook.
    class Program
    {
        static void Main(string[] args)
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            try
            {
                // Ensure the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Get the first worksheet (adjust index if needed)
                Worksheet worksheet = workbook.Worksheets[0];

                // Retrieve the textbox shape by name or fallback to the first shape
                Shape textBox = null;

                // Search for a shape named "TextBox 1"
                foreach (Shape shape in worksheet.Shapes)
                {
                    if (shape.Name.Equals("TextBox 1", StringComparison.OrdinalIgnoreCase))
                    {
                        textBox = shape;
                        break;
                    }
                }

                // If not found, use the first shape if any
                if (textBox == null && worksheet.Shapes.Count > 0)
                {
                    textBox = worksheet.Shapes[0];
                }

                if (textBox == null)
                {
                    Console.WriteLine("No textbox shape found in the worksheet.");
                }
                else
                {
                    // Apply left horizontal alignment to the text inside the textbox
                    textBox.TextHorizontalAlignment = TextAlignmentType.Left;
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                // Catch any runtime exceptions and display the message
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
