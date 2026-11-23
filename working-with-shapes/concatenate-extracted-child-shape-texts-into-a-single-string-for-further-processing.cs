// Title: Concatenate text from all shapes in an Excel workbook using Aspose.Cells for .NET
// Common Searches: C# code to read and combine text from every shape in an Excel file with Aspose.Cells | How to loop through worksheet shapes and get their Text property using Aspose.Cells | Aspose.Cells extract shape text from multiple worksheets and merge into a single string | Retrieve and concatenate child shape texts in a .xlsx workbook with the Aspose.Cells .NET API | Combine all shape captions from an Excel workbook programmatically in C#
// Tags: Aspose.Cells shape text extraction | concatenate shape texts C# | iterate worksheet shapes Aspose.Cells | StringBuilder aggregation Excel shape text | read shape Text property .NET

using System;
using System.IO;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an existing Excel workbook, iterates through each worksheet's Shapes collection, appends the Text property of every non‑empty shape to a StringBuilder, outputs the combined string, and saves the workbook.
class ShapeTextConcatenator
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists before attempting to load it
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file '{inputPath}' not found.");
            return;
        }

        Workbook workbook;
        try
        {
            workbook = new Workbook(inputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error loading workbook: {ex.Message}");
            return;
        }

        var concatenatedText = new StringBuilder();

        try
        {
            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                ShapeCollection shapes = sheet.Shapes;

                // Iterate through each shape on the worksheet
                foreach (Shape shape in shapes)
                {
                    // Append text from the shape (group shapes are ignored as they do not expose child shapes directly)
                    AppendShapeText(shape, concatenatedText);
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error processing shapes: {ex.Message}");
            return;
        }

        string result = concatenatedText.ToString();

        Console.WriteLine("Concatenated Text:");
        Console.WriteLine(result);

        // Save the workbook
        try
        {
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error saving workbook: {ex.Message}");
        }
    }

    // Helper method to append text from a shape if it contains any
    private static void AppendShapeText(Shape shape, StringBuilder sb)
    {
        // Shapes that contain text have a non‑empty Text property
        if (!string.IsNullOrEmpty(shape.Text))
        {
            sb.Append(shape.Text);
        }
    }
}
