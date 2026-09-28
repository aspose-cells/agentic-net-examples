// Title: Log each SmartArt shape’s type, row/column position, and dimensions from an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that loops through all worksheets, identifies SmartArt shapes, and prints their name, type, upper‑left row and column, width, and height. | Adjust the shape‑logging script to output only the bounding box (row, column, width, height) for SmartArt objects and gracefully handle missing worksheets. | Create a reusable C# method that receives a Worksheet and returns a collection of SmartArt shape details (type, position, size) using Aspose.Cells, with per‑shape exception handling.
// Common Searches: how to enumerate smartart shapes and get their coordinates with aspose.cells in c# | c# asp.net retrieve shape width and height from excel using aspose.cells | list smartart objects with row and column indices in an xlsx file using Aspose.Cells | asp.net core log details of each shape in workbook worksheets | asp.net get smartart shape type and dimensions from excel workbook
// Tags: Aspose.Cells enumerate SmartArt shapes C# | log shape dimensions Excel Aspose | retrieve shape position Aspose.Cells worksheet | SmartArt shape properties extraction .NET | per‑shape exception handling Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// The example loads an Excel workbook, iterates through every worksheet, and for each SmartArt shape prints the worksheet name, shape name, type, upper‑left row and column indices, and the shape's width and height in points, while handling errors on a per‑shape basis.
class Program
{
    static void Main()
    {
        string filePath = "input.xlsx";

        // Verify that the input file exists before attempting to load it
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"File not found: {filePath}");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(filePath);

            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through each shape on the worksheet
                foreach (Shape shape in sheet.Shapes)
                {
                    try
                    {
                        // Output basic shape information
                        Console.WriteLine($"Worksheet: {sheet.Name}");
                        Console.WriteLine($"Shape Name: {shape.Name}");
                        Console.WriteLine($"Shape Type: {shape.Type}");

                        // Position information (row and column indices)
                        int upperRow = shape.UpperLeftRow;
                        int upperColumn = shape.UpperLeftColumn;

                        // Size information (in points)
                        double width = shape.Width;
                        double height = shape.Height;

                        Console.WriteLine($"Position - Row: {upperRow}, Column: {upperColumn}");
                        Console.WriteLine($"Size - Width: {width} pts, Height: {height} pts");
                        Console.WriteLine(new string('-', 60));
                    }
                    catch (Exception shapeEx)
                    {
                        // Handle any errors that occur while processing an individual shape
                        Console.WriteLine($"Error processing shape '{shape.Name}': {shapeEx.Message}");
                    }
                }
            }
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
