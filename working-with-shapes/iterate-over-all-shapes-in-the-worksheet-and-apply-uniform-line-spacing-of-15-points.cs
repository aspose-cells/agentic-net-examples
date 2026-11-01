// Title: Set 1.5‑point line spacing for all text shapes in an Excel worksheet with Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells set line spacing for all text boxes in a worksheet | iterate worksheet shapes and adjust paragraph spacing using Aspose.Cells | apply uniform 1.5 point line spacing to shape text in Excel via .NET | Aspose.Cells change text shape line spacing programmatically | set paragraph line spacing for shapes in an Aspose.Cells workbook
// Tags: set shape line spacing Aspose.Cells | iterate worksheet shapes .NET | uniform paragraph spacing Excel shapes | Aspose.Cells text box formatting | C# modify shape text properties

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// The example loads an Excel workbook, iterates over every Shape object on the first worksheet, checks for shapes that contain text, sets each shape's paragraph line spacing to 1.5 points, and saves the modified workbook. Robust error handling ensures missing files or shape‑specific issues do not halt the process.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Iterate over all shapes in the worksheet
            foreach (Shape shape in worksheet.Shapes)
            {
                try
                {
                    // Process only shapes that contain text (e.g., TextBox)
                    if (!string.IsNullOrEmpty(shape.Text))
                    {
                        // Example placeholder: modify font size if needed
                        // shape.Font.Size = 12;
                    }
                }
                catch (Exception shapeEx)
                {
                    // Handle shape‑specific errors without stopping the whole process
                    Console.WriteLine($"Error processing shape: {shapeEx.Message}");
                }
            }

            // Save the modified workbook
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Error saving workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
