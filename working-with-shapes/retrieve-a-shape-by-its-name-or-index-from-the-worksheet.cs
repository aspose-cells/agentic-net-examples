// Title: Retrieve a Worksheet Shape by Index or Name with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to add a rectangle and an oval to a worksheet, access the ShapeCollection, and fetch a shape either by its zero‑based index or by its default name (e.g., "Rectangle 1" or "Oval 1"). The sample prints key properties and saves the workbook as RetrieveShapeDemo.xlsx.
// Keywords: Aspose.Cells shape by index | Aspose.Cells get shape by name | C# retrieve worksheet shape | Aspose.Cells ShapeCollection example | access shapes in Excel with Aspose
// Common Searches: Aspose.Cells retrieve shape by index | How to get a shape by name in Aspose.Cells C# | ShapeCollection zero based index Aspose.Cells | Default shape names Aspose.Cells worksheet
// Developer Intent: Find the simplest way to locate a specific shape in an Excel worksheet using Aspose.Cells, either through its numeric index or its assigned name.
// Use Cases: Validate that a newly added shape exists by reading its type and name via index. | Check for a known shape name before modifying its size, color, or position. | Loop through all shapes, retrieve each by name, and apply conditional formatting based on shape type.
// AI Prompts: Generate code to change a shape's fill color after retrieving it by name with Aspose.Cells for .NET. | Show robust error handling when accessing a shape by index, including out‑of‑range checks. | Explain how to rename a shape after obtaining it from the worksheet's ShapeCollection.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // Demonstrates how to add a rectangle and an oval to a worksheet, access the ShapeCollection, and fetch a shape either by its zero‑based index or by its default name (e.g., "Rectangle 1" or "Oval 1"). The sample prints key properties and saves the workbook as RetrieveShapeDemo.xlsx.
    public class RetrieveShapeDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add shapes to the worksheet
                // Rectangle will have default name "Rectangle 1"
                worksheet.Shapes.AddRectangle(2, 0, 2, 0, 100, 200);
                // Oval will have default name "Oval 1"
                worksheet.Shapes.AddOval(5, 0, 5, 0, 80, 120);

                // Get the shape collection
                ShapeCollection shapes = worksheet.Shapes;

                // Retrieve shape by index (zero‑based)
                Shape shapeByIndex = shapes[0]; // first shape added
                Console.WriteLine($"Shape at index 0: Type={shapeByIndex.Type}, Name={shapeByIndex.Name}");

                // Retrieve shape by name
                Shape shapeByName = shapes["Rectangle 1"];
                if (shapeByName != null)
                {
                    Console.WriteLine($"Shape with name 'Rectangle 1': Type={shapeByName.Type}, Id={shapeByName.Id}");
                }

                // Retrieve the second shape by its default name
                Shape ovalShape = shapes["Oval 1"];
                if (ovalShape != null)
                {
                    Console.WriteLine($"Oval shape: Width={ovalShape.Width}, Height={ovalShape.Height}");
                }

                // Save the workbook
                workbook.Save("RetrieveShapeDemo.xlsx");
                Console.WriteLine("Workbook saved as RetrieveShapeDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            RetrieveShapeDemo.Run();
        }
    }
}
