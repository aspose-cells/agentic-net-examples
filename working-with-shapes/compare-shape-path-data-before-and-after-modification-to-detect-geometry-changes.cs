// Title: Detect worksheet shape geometry changes by comparing width and height before and after resizing with Aspose.Cells for .NET
// AI Prompts: Generate C# code that adds a rectangle shape to a worksheet, records its initial Width and Height, resizes the shape, and returns a boolean indicating whether the geometry changed using Aspose.Cells. | Show how to capture a shape's original dimensions, modify them, and compare the before/after values to detect size alteration in Aspose.Cells for .NET. | Provide a snippet that logs the original and updated geometry of a worksheet shape and prints true if the dimensions differ, using the Aspose.Cells API.
// Common Searches: Aspose.Cells C# compare shape dimensions before and after resize | how to check if a worksheet shape size was changed programmatically Aspose.Cells | detect geometry modification of rectangle shape in Excel using Aspose.Cells .NET | C# code to determine if shape width or height has been altered in Aspose.Cells workbook
// Tags: shape dimension comparison Aspose.Cells | detect geometry change worksheet shape | rectangle shape resize tracking Aspose.Cells | shape width height verification C# | Aspose.Cells shape geometry monitoring

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The program creates a workbook, adds a rectangle shape, records its initial Width and Height, changes the dimensions, compares the original and modified geometry strings, outputs whether the geometry changed, and saves the workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet
            // Parameters: shape type, upper left row, upper left column, top offset, left offset, width, height
            Shape shape = worksheet.Shapes.AddShape(MsoDrawingType.Rectangle, 1, 1, 0, 0, 100, 50);

            // Capture the original geometry data (width and height)
            string originalGeometry = $"{shape.Width},{shape.Height}";

            // Modify the shape's geometry (change width and height)
            shape.Width = 200;   // new width
            shape.Height = 100;  // new height

            // Capture the geometry data after modification
            string modifiedGeometry = $"{shape.Width},{shape.Height}";

            // Compare the two geometry strings to detect changes
            bool geometryChanged = !originalGeometry.Equals(modifiedGeometry, StringComparison.Ordinal);

            // Output the comparison results
            Console.WriteLine("Original Geometry (W,H):  " + originalGeometry);
            Console.WriteLine("Modified Geometry (W,H):  " + modifiedGeometry);
            Console.WriteLine("Geometry changed:          " + geometryChanged);

            // Save the workbook (optional, demonstrates lifecycle usage)
            string outputFile = "ShapePathComparison.xlsx";

            // Ensure the directory exists before saving
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputFile));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(outputFile);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputFile)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
