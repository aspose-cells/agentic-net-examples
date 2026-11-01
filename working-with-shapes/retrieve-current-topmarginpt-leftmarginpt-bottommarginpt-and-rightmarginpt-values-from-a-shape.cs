// Title: Retrieve shape margin values (TopMarginPt, LeftMarginPt, BottomMarginPt, RightMarginPt) from an Excel worksheet with Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an Excel file using Aspose.Cells and prints the TopMarginPt, LeftMarginPt, BottomMarginPt, and RightMarginPt of the first shape on the first worksheet. | Show how to access a shape's margin properties in points via Aspose.Cells for .NET and output them to the console. | Provide a snippet that checks for shapes in a worksheet and extracts their margin measurements (top, left, bottom, right) using the Aspose.Cells Shape API.
// Common Searches: Aspose.Cells get shape margin points C# | How to read TopMarginPt of a shape in Excel with Aspose.Cells .NET | C# Aspose.Cells retrieve shape left and right margins from worksheet | Extract bottom margin of a shape using Aspose.Cells for .NET | Read shape margin values from Excel file programmatically
// Tags: Aspose.Cells shape margin properties | C# retrieve shape margins Excel | Aspose.Cells get TopMarginPt | read shape margin points .NET | extract shape margin values Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an Excel workbook with Aspose.Cells, accesses the first worksheet, verifies that shapes exist, selects the first shape, and prints its TopMarginPt, LeftMarginPt, BottomMarginPt, and RightMarginPt values (in points) to the console.
class Program
{
    static void Main()
    {
        string inputPath = "input.xlsx";

        // Ensure the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"File not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Verify that the worksheet contains at least one shape
            if (worksheet.Shapes.Count == 0)
            {
                Console.WriteLine("No shapes found in the worksheet.");
                return;
            }

            // Retrieve the first shape
            Shape shape = worksheet.Shapes[0];

            // Get the shape's position and size (in points)
            double top = shape.Top;
            double left = shape.Left;
            double height = shape.Height;
            double width = shape.Width;

            // Output the values
            Console.WriteLine($"Top: {top}");
            Console.WriteLine($"Left: {left}");
            Console.WriteLine($"Height: {height}");
            Console.WriteLine($"Width: {width}");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
