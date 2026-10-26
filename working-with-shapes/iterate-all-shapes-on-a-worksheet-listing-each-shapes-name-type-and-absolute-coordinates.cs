// Title: C# example: enumerate all shapes in an Excel worksheet with Aspose.Cells and obtain each shape's name, type and absolute coordinates
// Common Searches: how to enumerate shapes in an Excel file using Aspose.Cells C# | retrieve shape name and position from a worksheet with Aspose.Cells | list all drawing objects coordinates Aspose.Cells C# example | Aspose.Cells get shape dimensions (left, top, width, height) in C# | C# code to iterate worksheet shapes and output their properties
// Tags: worksheet shape listing Aspose.Cells C# | shape absolute coordinates Aspose.Cells | extract shape type and dimensions Excel C# | Aspose.Cells shape enumeration sample | drawing objects position retrieval Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample loads an existing Excel workbook, accesses a worksheet, iterates through every Shape object on that sheet, and prints each shape's Name, Type, Left, Top, Width, and Height values before optionally saving the workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (change index as needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Iterate through all shapes on the worksheet
            foreach (Shape shape in sheet.Shapes)
            {
                // Shape name
                string name = shape.Name;

                // Shape type (enum value)
                string type = shape.Type.ToString();

                // Absolute coordinates (points)
                double left = shape.Left;   // X position of upper‑left corner
                double top = shape.Top;     // Y position of upper‑left corner
                double width = shape.Width; // Width of the shape
                double height = shape.Height; // Height of the shape

                Console.WriteLine($"Name: {name}, Type: {type}, Left: {left}, Top: {top}, Width: {width}, Height: {height}");
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook (optional, unchanged)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
