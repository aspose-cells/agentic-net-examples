// Title: How to locate and verify a named shape in an Excel worksheet using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an Excel workbook with Aspose.Cells, searches the first worksheet's Shapes collection for a shape named "MyShape", and prints its Type when found. | Provide a C# example that iterates over Worksheet.Shapes to determine whether a specific shape exists and handles the case where the shape is missing.
// Common Searches: Aspose.Cells C# find shape by its Name property in a worksheet | C# code to check if a shape exists in an Excel file using Aspose.Cells | How to get the Type of a named shape in Aspose.Cells for .NET | Iterate through worksheet shapes collection to locate a specific shape Aspose.Cells
// Tags: shape name lookup Aspose.Cells | verify shape presence worksheet C# | retrieve shape type Aspose.Cells | enumerate worksheet shapes Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// Loads an Excel workbook, iterates through the first worksheet's Shapes collection to find a shape named "MyShape", and outputs whether the shape exists together with its Type.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the input workbook
            string inputPath = "input.xlsx";

            // Verify that the file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (index 0)
            Worksheet worksheet = workbook.Worksheets[0];

            // Name of the shape to retrieve
            string targetShapeName = "MyShape";

            // Search for the shape by its name
            Shape foundShape = null;
            foreach (Shape shape in worksheet.Shapes)
            {
                if (shape.Name == targetShapeName)
                {
                    foundShape = shape;
                    break;
                }
            }

            // Output the result
            if (foundShape != null)
            {
                Console.WriteLine($"Shape '{targetShapeName}' exists. Type: {foundShape.Type}");
            }
            else
            {
                Console.WriteLine($"Shape '{targetShapeName}' does not exist in the worksheet.");
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
