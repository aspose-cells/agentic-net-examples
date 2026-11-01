// Title: Remove all shapes from an Excel worksheet and add a rectangular shape with custom text using Aspose.Cells for .NET
// Common Searches: C# code to delete every shape in an Aspose.Cells worksheet | how to add a rectangle shape with custom text to an existing Excel file using Aspose.Cells | sample for clearing all shapes before inserting a new shape in a .NET workbook | Aspose.Cells example for applying a tiled texture to a shape in C# | remove shapes from first worksheet and save workbook with Aspose.Cells
// Tags: remove worksheet shapes Aspose.Cells | add rectangle shape with text Aspose.Cells | clear existing shapes before new shape .NET | shape fill texture tile Aspose.Cells | load workbook modify shapes Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The C# program verifies the input Excel file, loads it with Aspose.Cells, accesses the first worksheet, removes all existing shapes by iterating backwards, adds a new rectangular shape at row 2 column 2 with specified height and width, sets the shape's Text property to "Custom Text", notes that texture properties are unavailable in the current version, saves the workbook to an output file, and handles any exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input workbook exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Remove all existing shapes from the worksheet
            // Iterate backwards to avoid index shifting when removing
            for (int i = sheet.Shapes.Count - 1; i >= 0; i--)
            {
                sheet.Shapes.RemoveAt(i);
            }

            // Add a new rectangular shape
            // Parameters: drawing type, upper left row, upper left column, top offset, left offset, height, width
            Shape newShape = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle, // Shape type
                2,                        // Upper left row (zero‑based)
                2,                        // Upper left column (zero‑based)
                0,                        // Top offset in pixels
                0,                        // Left offset in pixels
                100,                      // Height in points
                200);                     // Width in points

            // Set custom text inside the shape
            newShape.Text = "Custom Text";

            // NOTE: The current Aspose.Cells version used in this project does not expose
            // texture‑related members (Texture, TextureAlignment) on FillFormat.
            // If a newer version is referenced, the following lines can be uncommented
            // and adjusted accordingly to apply a tiled texture:
            // newShape.Fill.Texture = TextureType.Texture1;
            // newShape.Fill.TextureAlignment = TextureAlignmentType.Tile;

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
