// Title: How to link each shape on an Excel worksheet to a specific cell so the shape text updates dynamically using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that loads an existing workbook, iterates over all shapes on the first worksheet, sets each shape's Text property to a formula referencing sequential cells (A1, A2, …), and configures the shape placement to MoveAndSize before saving the file. | Write a C# program with Aspose.Cells that binds every shape to a matching cell so that any change in the cell value automatically updates the shape's displayed text, then export the modified workbook.
// Common Searches: Aspose.Cells C# link shape text to a cell value | set shape formula to reference a cell using Aspose.Cells .NET | make Excel shapes move and size with their cells in Aspose.Cells | iterate over shapes collection and bind each to a cell Aspose.Cells | assign sequential cell references to shapes in an Excel workbook with Aspose.Cells
// Tags: link shape to cell Aspose.Cells | shape text formula Aspose.Cells C# | move and size placement Aspose.Cells | iterate shapes collection Aspose.Cells | bind Excel shape to cell value .NET | sequential cell reference shapes Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads a workbook, loops through every shape on the first worksheet, assigns each shape a formula that points to a sequential cell (A1, A2, …), sets the shape's placement to MoveAndSize so it follows the cell, and saves the updated file.
class ShapeLinkExample
{
    static void Main()
    {
        const string inputPath = "Input.xlsx";
        const string outputPath = "Output.xlsx";

        // Verify that the input workbook exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file not found at '{inputPath}'.");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Get the collection of shapes on the worksheet
            ShapeCollection shapes = sheet.Shapes;

            // Loop through each shape and link it to a corresponding cell (A1, A2, A3, ...)
            for (int i = 0; i < shapes.Count; i++)
            {
                Shape shape = shapes[i];

                // Determine the target cell address (e.g., A1, A2, ...)
                string cellAddress = $"A{i + 1}";

                // Set the shape's text to a formula that references the target cell
                shape.Text = $"={cellAddress}";

                // Ensure the shape moves and sizes with the cell
                shape.Placement = PlacementType.MoveAndSize;
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
