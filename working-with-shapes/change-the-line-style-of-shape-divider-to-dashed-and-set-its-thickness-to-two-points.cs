// Title: Change the line style of the "Divider" shape to dashed and set its thickness to 2 points using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells change rectangle shape line to dashed | How to set a shape's line thickness to 2 points with Aspose.Cells | Aspose.Cells find shape by name and modify its line style | Set dash style for Excel shape named Divider using Aspose.Cells .NET | Update line weight of a shape in an existing workbook with Aspose.Cells
// Tags: set shape dash style Aspose.Cells | shape line weight Aspose.Cells .NET | modify rectangle line format Excel C# | MsoLineDashStyle dash Aspose.Cells | find shape by name Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace ShapeLineStyleDemo
{
    // The example loads or creates an Excel workbook, locates a rectangle shape named "Divider", and updates its LineFormat by applying a dashed dash style and setting the line weight to 2 points before saving the file.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                Workbook workbook;

                // Load existing workbook or create a new one if the file does not exist
                if (File.Exists(inputPath))
                {
                    workbook = new Workbook(inputPath);
                }
                else
                {
                    workbook = new Workbook();
                    Worksheet ws = workbook.Worksheets[0];

                    // Add a simple rectangle shape and name it "Divider"
                    Shape shape = ws.Shapes.AddRectangle(1, 0, 1, 0, 200, 20);
                    shape.Name = "Divider";
                }

                Worksheet worksheet = workbook.Worksheets[0];

                // Find the shape named "Divider"
                Shape dividerShape = null;
                foreach (Shape shape in worksheet.Shapes)
                {
                    if (string.Equals(shape.Name, "Divider", StringComparison.OrdinalIgnoreCase))
                    {
                        dividerShape = shape;
                        break;
                    }
                }

                if (dividerShape == null)
                {
                    Console.WriteLine("Shape named 'Divider' was not found in the worksheet.");
                    return;
                }

                // Access the line formatting of the shape
                LineFormat line = dividerShape.Line;

                // Set dash style to dashed and line thickness to 2 points
                line.DashStyle = MsoLineDashStyle.Dash;
                line.Weight = 2.0f; // thickness in points

                // Optional: set line color for visibility (commented out if not supported)
                // line.Color = Color.Black;

                // Save the modified workbook
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Shape 'Divider' updated successfully. Saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
