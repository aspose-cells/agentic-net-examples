// Title: How to retrieve a named shape from an Excel worksheet and verify its existence using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that loads an Excel workbook with Aspose.Cells, accesses a worksheet, and uses the Shapes collection indexer to fetch a shape by its name, then prints whether the shape was found and its type. | Show how to handle the scenario where the requested shape name is not present in the worksheet when working with Aspose.Cells in a .NET application.
// Common Searches: Aspose.Cells C# retrieve shape by its name from worksheet | Check if a drawing object exists in an Excel sheet using Aspose.Cells | How to get shape type after finding a named shape with Aspose.Cells .NET | Worksheet.Shapes indexer example for named shapes in Aspose.Cells
// Tags: Aspose.Cells retrieve shape by name | C# worksheet shapes indexer | verify shape existence Excel | Aspose.Cells shape type detection | handle missing shape Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// // Loads 'input.xlsx', accesses the first worksheet, attempts to retrieve a shape named 'MyShape' via worksheet.Shapes["MyShape"], and prints whether the shape exists along with its ShapeType.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook from the file
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Name of the shape to retrieve
            string targetShapeName = "MyShape";

            // Retrieve the shape by its name
            Shape shape = worksheet.Shapes[targetShapeName];

            // Check if the shape exists and output its type
            if (shape != null)
            {
                Console.WriteLine($"Shape '{targetShapeName}' exists. Type: {shape.Type}");
            }
            else
            {
                Console.WriteLine($"Shape '{targetShapeName}' does not exist.");
            }
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
