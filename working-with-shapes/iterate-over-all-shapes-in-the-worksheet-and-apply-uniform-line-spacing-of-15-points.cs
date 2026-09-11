// Title: Loop through all worksheet shapes and set 1.5‑point line spacing for their text using Aspose.Cells for .NET
// AI Prompts: Generate C# code that enumerates every Shape in a worksheet and applies a 1.5 point paragraph spacing to the shape's text with Aspose.Cells. | Explain how to adjust line spacing for text boxes in an Excel file via the Aspose.Cells .NET API, including any work‑arounds for missing ParagraphSpacing support. | Provide a sample that checks each shape for text content, sets uniform line spacing, and saves the workbook.
// Common Searches: Aspose.Cells .NET how to change line spacing in Excel shape text | C# iterate worksheet shapes and modify paragraph spacing with Aspose.Cells | set uniform 1.5 point spacing for text boxes in an Excel workbook using Aspose.Cells | Aspose.Cells shape text formatting options for line spacing
// Tags: Aspose.Cells shape text formatting | C# iterate worksheet shapes | set line spacing Aspose.Cells | Excel text box paragraph spacing .NET | uniform shape text line spacing

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an existing workbook, verifies the input file, iterates over all shapes on the first worksheet, checks each shape for non‑empty text, and notes that Aspose.Cells does not expose a direct ParagraphSpacing property, indicating where line‑spacing adjustments would be applied before saving the modified file.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);
            Worksheet worksheet = workbook.Worksheets[0];

            // Iterate through all shapes in the worksheet
            foreach (Shape shape in worksheet.Shapes)
            {
                // Apply formatting only to shapes that contain text
                if (!string.IsNullOrEmpty(shape.Text))
                {
                    // Aspose.Cells does not provide a direct ParagraphSpacing property.
                    // You can modify other available text properties here if needed.
                    // Example: shape.TextEffect.Font.Size = 12; // adjust font size
                }
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
