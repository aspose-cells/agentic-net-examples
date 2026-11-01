// Title: Concatenate text from all worksheet shapes into a single string with Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an Excel workbook using Aspose.Cells, iterates over every Shape on the first worksheet, and appends each Shape.Text to a StringBuilder to produce one combined string. | Add error handling so that shapes without text or shapes that raise exceptions are skipped while the concatenation continues uninterrupted. | Modify the example to write the final concatenated shape text to a UTF‑8 file instead of outputting it to the console.
// Common Searches: aspocells .net read all shape texts from an Excel worksheet | how to merge captions of multiple shapes in an Excel file using C# | C# Aspose.Cells concatenate shape.Text values from a worksheet | skip null shape text when iterating Excel shapes with Aspose.Cells | save concatenated shape text to a file using Aspose.Cells in C#
// Tags: concatenate shape text Aspose.Cells | iterate worksheet shapes .NET | extract shape captions Excel | shape text aggregation C# | Aspose.Cells read Shape.Text property

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;
using System.Text;

// The program loads 'input.xlsx', loops through every Shape on the first worksheet, appends each non‑empty Shape.Text to a StringBuilder, and outputs the combined string, with error handling for missing files and shape‑specific exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input workbook exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file '{inputPath}' not found.");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);
            Worksheet sheet = workbook.Worksheets[0];

            StringBuilder sb = new StringBuilder();

            // Iterate through all shapes on the worksheet
            foreach (Shape shape in sheet.Shapes)
            {
                try
                {
                    // Append text from the shape if it exists
                    if (!string.IsNullOrEmpty(shape.Text))
                    {
                        sb.Append(shape.Text);
                    }
                }
                catch (Exception innerEx)
                {
                    // Log shape‑specific errors but continue processing other shapes
                    Console.WriteLine($"Error processing a shape: {innerEx.Message}");
                }
            }

            // Output the concatenated text
            string concatenatedText = sb.ToString();
            Console.WriteLine(concatenatedText);
        }
        catch (Exception ex)
        {
            // Report any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
