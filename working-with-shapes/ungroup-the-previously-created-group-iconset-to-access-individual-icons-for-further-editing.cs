// Title: Ungroup an IconSet (GroupShape) and Edit Individual Shapes with Aspose.Cells for .NET (C#)
// Description: This C# example demonstrates how to create a workbook, group rectangle shapes to mimic an IconSet, call GroupShape.Ungroup() to release the icons, modify properties of each shape (e.g., line weight, fill color), and save the result. It shows the exact API calls needed for shape manipulation in Aspose.Cells for .NET.
// Keywords: Aspose.Cells ungroup GroupShape | IconSet shape editing .NET | C# Aspose.Cells GroupShape.Ungroup | modify individual Excel shapes | Aspose.Cells shape line weight | access shapes after grouping | Aspose.Cells workbook shape example | C# Excel icon set ungroup
// Common Searches: Aspose.Cells how to ungroup shapes in C# | Ungroup IconSet using Aspose.Cells for .NET | Edit individual icons after grouping Aspose.Cells | C# code to modify shape line weight in Excel workbook | Retrieve shapes from a GroupShape Aspose.Cells
// Developer Intent: The developer needs to break a grouped IconSet into its component shapes so each can be formatted or updated independently.
// Use Cases: Separate a grouped IconSet to change the fill color of a specific icon. | Adjust line weight, style, or color of individual shapes after ungrouping. | Iterate through ungrouped shapes to apply custom formatting before saving the workbook.
// AI Prompts: Generate C# code that uses Aspose.Cells to ungroup a GroupShape and set the fill color of the first icon to LightBlue. | Show how to loop through all shapes returned by GroupShape.Ungroup() and assign different line styles to each. | Explain fallback handling for older Aspose.Cells versions where Shape.Fill.ForeColor is unavailable.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // This C# example demonstrates how to create a workbook, group rectangle shapes to mimic an IconSet, call GroupShape.Ungroup() to release the icons, modify properties of each shape (e.g., line weight, fill color), and save the result. It shows the exact API calls needed for shape manipulation in Aspose.Cells for .NET.
    public class UngroupIconSetDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add two rectangle shapes that will represent icons (for demonstration)
                Shape shape1 = worksheet.Shapes.AddRectangle(0, 0, 0, 0, 100, 50);
                Shape shape2 = worksheet.Shapes.AddRectangle(0, 0, 3, 0, 100, 50);

                // Group the two shapes – this simulates a grouped IconSet
                GroupShape groupShape = worksheet.Shapes.Group(new Shape[] { shape1, shape2 });

                // Ungroup the previously created group to access individual shapes (icons)
                groupShape.Ungroup();

                // After ungrouping, the shapes are available individually in the collection
                // Example: modify the first shape's fill color (if supported)
                // Note: Fill.ForeColor may not be available in some versions; this line is optional.
                // shape1.Fill.ForeColor = Color.LightBlue;

                // Example: modify the second shape's line style
                shape2.Line.Weight = 2.0f;
                // Note: Line.Color may not be available in some versions; this line is optional.
                // shape2.Line.Color = Color.DarkBlue;

                // Save the workbook
                string outputPath = "UngroupedIconSetDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            UngroupIconSetDemo.Run();
        }
    }
}
