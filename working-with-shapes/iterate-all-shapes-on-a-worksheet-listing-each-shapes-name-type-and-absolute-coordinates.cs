// Title: Iterate all shapes on an Excel worksheet and list each shape's name, type, and absolute row/column coordinates using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells list all shapes in a worksheet with their positions | how to get shape name, type and cell coordinates with Aspose.Cells .NET | iterate ShapeCollection to retrieve row and column indexes in Aspose.Cells | Aspose.Cells example for reading shape dimensions and location in C#
// Tags: iterate worksheet shape collection Aspose.Cells | retrieve shape name and type Aspose.Cells | get shape absolute row column indexes Aspose.Cells | read shape dimensions pixels Aspose.Cells | Aspose.Cells shape enumeration .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample loads an Excel workbook, accesses the first worksheet, and loops through its ShapeCollection. For each shape it determines a fallback name, obtains the runtime type, reads the UpperLeftRow and UpperLeftColumn indexes (absolute coordinates), and reports the width and height in pixels. Results are printed to the console, with error handling for missing files and individual shape processing failures.
class Program
{
    static void Main(string[] args)
    {
        // Path to the Excel file.
        string filePath = "input.xlsx";

        // Verify that the input file exists.
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"Error: The file \"{filePath}\" does not exist.");
            return;
        }

        try
        {
            // Load the workbook.
            Workbook workbook = new Workbook(filePath);

            // Use the first worksheet.
            Worksheet worksheet = workbook.Worksheets[0];

            // Get the collection of shapes on the worksheet.
            ShapeCollection shapes = worksheet.Shapes;

            Console.WriteLine($"Worksheet \"{worksheet.Name}\" contains {shapes.Count} shape(s).");
            Console.WriteLine();

            for (int i = 0; i < shapes.Count; i++)
            {
                try
                {
                    Shape shape = shapes[i];

                    // Shape name (fallback if empty).
                    string shapeName = string.IsNullOrEmpty(shape.Name) ? $"Shape_{i}" : shape.Name;

                    // Determine shape type using runtime type name (compatible with all versions).
                    string shapeType = shape.GetType().Name;

                    // Absolute position (zero‑based row/column indexes).
                    int upperLeftRow = shape.UpperLeftRow;
                    int upperLeftColumn = shape.UpperLeftColumn;

                    // Size in pixels (cast to int for display).
                    int widthPixels = (int)shape.Width;
                    int heightPixels = (int)shape.Height;

                    Console.WriteLine($"Shape {i + 1}:");
                    Console.WriteLine($"  Name     : {shapeName}");
                    Console.WriteLine($"  Type     : {shapeType}");
                    Console.WriteLine($"  Position : Row {upperLeftRow}, Column {upperLeftColumn}");
                    Console.WriteLine($"  Size     : Width {widthPixels}px, Height {heightPixels}px");
                    Console.WriteLine();
                }
                catch (Exception shapeEx)
                {
                    Console.WriteLine($"Error processing shape at index {i}: {shapeEx.Message}");
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("An exception occurred while processing the workbook:");
            Console.WriteLine(ex.Message);
        }
    }
}
