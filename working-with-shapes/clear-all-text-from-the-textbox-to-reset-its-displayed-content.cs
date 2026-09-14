// Title: How to clear text from all textbox shapes in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that opens an existing .xlsx file with Aspose.Cells, iterates over each Shape on a worksheet, sets the Shape.Text property to an empty string for text‑enabled shapes, and saves the modified workbook. | Create a robust Aspose.Cells example that removes the displayed content of every textbox shape, includes try‑catch handling for shapes that do not support text, and writes the result to a specified output path.
// Common Searches: Aspose.Cells C# clear textbox shape text in Excel file | remove all textbox content from worksheet using Aspose.Cells .NET | iterate over worksheet shapes and reset Text property with Aspose.Cells | how to delete text inside Excel shapes programmatically in C#
// Tags: empty shape Text property Aspose.Cells | iterate worksheet shapes Aspose.Cells | reset shape text .xlsx C# | skip non‑text shapes Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // The program loads an existing Excel workbook, loops through every shape on the first worksheet, empties the Text property of shapes that support text (ignoring others), and saves the cleaned workbook to a new file.
    class Program
    {
        static void Main(string[] args)
        {
            // Define input and output file paths
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            try
            {
                // Verify that the input workbook exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet (adjust index if needed)
                Worksheet worksheet = workbook.Worksheets[0];

                // Iterate through all shapes on the worksheet
                foreach (Shape shape in worksheet.Shapes)
                {
                    try
                    {
                        // Clear the displayed text for any shape that supports text
                        shape.Text = string.Empty;
                    }
                    catch (Exception shapeEx)
                    {
                        // Log shape-specific errors without stopping the whole process
                        Console.WriteLine($"Error processing shape: {shapeEx.Message}");
                    }
                }

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook with the cleared textbox content
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
