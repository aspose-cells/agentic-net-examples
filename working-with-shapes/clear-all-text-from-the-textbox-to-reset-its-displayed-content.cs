// Title: Clear text from every textbox shape in an Excel worksheet with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code using Aspose.Cells that loads a workbook, iterates over all shapes on the first worksheet, sets the Text property of each textbox shape to an empty string, and saves the file. | Create a reusable method in C# that accepts input and output file paths, clears the content of all textbox‑type shapes in the workbook using Aspose.Cells, and returns a success status. | Adapt the example to skip non‑textbox shapes while clearing text, and log the names of processed shapes, using Aspose.Cells for .NET.
// Common Searches: aspocells c# clear textbox shape text in excel file | remove all text from Excel textbox shapes using Aspose.Cells .NET | how to reset content of shape objects in an Aspose.Cells workbook | iterate over worksheet shapes and empty their Text property with Aspose.Cells | C# Aspose.Cells clear textbox content without affecting other shapes
// Tags: clear textbox shape text Aspose.Cells | iterate worksheet shapes C# | Aspose.Cells shape text reset | Excel workbook modify shape content .NET | handle non‑text shapes Aspose.Cells | save workbook after shape text removal

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an existing Excel file, loops through every shape on the first worksheet, attempts to set each shape's Text property to an empty string while safely ignoring shapes that do not support text, and then saves the modified workbook to a new file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            var workbook = new Workbook(inputPath);

            // Access the first worksheet
            var sheet = workbook.Worksheets[0];

            // Clear text in all textbox-like shapes
            foreach (Shape shape in sheet.Shapes)
            {
                try
                {
                    // Attempt to clear text; non‑text shapes will throw, which we ignore
                    shape.Text = string.Empty;
                }
                catch
                {
                    // Ignore shapes that do not support text
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
