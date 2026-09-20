// Title: Retrieve and display absolute positions of all shapes on an Excel worksheet using Aspose.Cells for .NET
// AI Prompts: Generate a C# console program that loads an Excel workbook with Aspose.Cells, iterates through Worksheet.Shapes, and prints each shape's Name, Top, Left, Width, and Height in points. | Write a C# method that takes a file path, validates the Excel file, and uses Aspose.Cells to output the absolute coordinates of every shape on the first worksheet. | Create an example that extends shape enumeration to include the shape type (e.g., picture, chart) together with its absolute position using the Aspose.Cells API.
// Common Searches: Aspose.Cells C# get absolute position of shapes in Excel worksheet | How to list shape top left coordinates with Aspose.Cells .NET | C# enumerate all shapes in an Excel file and display their dimensions using Aspose.Cells | Retrieve shape size and location from the first worksheet using Aspose.Cells library | Console application to print shape positions from Excel using Aspose.Cells for .NET
// Tags: Aspose.Cells retrieve shape coordinates | C# enumerate worksheet shapes | Aspose.Cells shape absolute position | Excel shape dimensions Aspose.Cells | Console output shape positions .NET

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// Loads a workbook, iterates through the Shapes collection of the first worksheet, and writes each shape's Name, Top, Left, Width, and Height (in points) to the console.
public class ShapePositionAnalyzer
{
    // Retrieves and displays absolute positions of all shapes on the first worksheet
    public static void DisplayShapePositions(string workbookPath)
    {
        if (!File.Exists(workbookPath))
        {
            Console.WriteLine($"Error: File not found - {workbookPath}");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(workbookPath);

            // Access the first worksheet (change index or name as needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Iterate through each shape in the worksheet
            foreach (Shape shape in sheet.Shapes)
            {
                // Absolute position values (in points)
                double top = shape.Top;      // distance from the top edge of the worksheet
                double left = shape.Left;    // distance from the left edge of the worksheet
                double width = shape.Width;
                double height = shape.Height;

                // Output shape details
                Console.WriteLine($"Shape Name: {shape.Name}");
                Console.WriteLine($"  Top: {top} pt, Left: {left} pt");
                Console.WriteLine($"  Width: {width} pt, Height: {height} pt");
                Console.WriteLine();
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred while processing the workbook: {ex.Message}");
        }
    }

    // Example entry point
    public static void Main(string[] args)
    {
        if (args.Length == 0)
        {
            Console.WriteLine("Usage: ShapePositionAnalyzer <excel-file-path>");
            return;
        }

        string filePath = args[0];
        DisplayShapePositions(filePath);
    }
}
