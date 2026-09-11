// Title: Detect changes in an Excel shape's geometry by comparing path data before and after resizing with Aspose.Cells for .NET
// AI Prompts: Generate C# code that records a shape's left, top, width, and height as a formatted string, resizes the shape, and then checks whether the geometry string has changed using Aspose.Cells. | Show how to retrieve a shape's geometry, modify its dimensions, and programmatically determine if the shape's path data differs after scaling with Aspose.Cells.
// Common Searches: how to programmatically compare Excel shape dimensions before and after scaling using Aspose.Cells C# | Aspose.Cells detect if shape size was altered in a worksheet | C# get shape geometry string Aspose.Cells and compare for changes | compare original and modified shape path data in Aspose.Cells workbook
// Tags: shape geometry comparison Aspose.Cells | detect shape size change .NET | retrieve shape dimensions Excel | compare original and modified shape data C# | shape scaling detection Aspose.Cells

using System;
using System.IO;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads or creates an Excel workbook, ensures a rectangle shape exists, captures its X, Y, width, and height as a string, enlarges the shape by 20 %, captures the new geometry string, compares the two strings to detect any geometry change, outputs the result, and saves the workbook.
class ShapePathComparison
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Load existing workbook or create a new one if the file is missing.
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
                workbook.Worksheets[0].Name = "Sheet1";
                workbook.Save(inputPath);
            }

            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure at least one shape exists; add a rectangle if none are present.
            if (worksheet.Shapes.Count == 0)
            {
                Shape rect = worksheet.Shapes.AddShape(MsoDrawingType.Rectangle, 1, 1, 0, 0, 100, 50);
                rect.Name = "SampleRectangle";
            }

            // Get the first shape on the worksheet.
            Shape shape = worksheet.Shapes[0];

            // Capture original geometry data.
            string originalData = GetShapeGeometryString(shape);

            // Modify the shape geometry (example: scale width and height).
            shape.Width = (int)(shape.Width * 1.2);
            shape.Height = (int)(shape.Height * 1.2);

            // Capture modified geometry data.
            string modifiedData = GetShapeGeometryString(shape);

            // Detect geometry change by comparing the data strings.
            bool geometryChanged = !originalData.Equals(modifiedData, StringComparison.Ordinal);

            // Output results.
            Console.WriteLine("Original Geometry: " + originalData);
            Console.WriteLine("Modified Geometry: " + modifiedData);
            Console.WriteLine("Geometry changed: " + geometryChanged);

            // Save the workbook.
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }

    // Helper method to create a simple string representation of a shape's geometry.
    static string GetShapeGeometryString(Shape shape)
    {
        try
        {
            // Use width, height, and position as a lightweight representation.
            StringBuilder sb = new StringBuilder();
            sb.AppendFormat("X={0},Y={1},Width={2},Height={3}", shape.Left, shape.Top, shape.Width, shape.Height);
            return sb.ToString();
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error retrieving shape geometry: " + ex.Message);
            return string.Empty;
        }
    }
}
