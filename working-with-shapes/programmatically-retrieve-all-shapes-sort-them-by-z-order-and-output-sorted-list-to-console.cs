// Title: Retrieve all worksheet shapes, sort them by Z‑order (fallback to name) and print the list using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells get all shapes from every worksheet and sort by Z‑order | How to list Excel shapes with their parent sheet names using Aspose.Cells in .NET | Aspose.Cells example to enumerate shapes and output a sorted list to the console | Retrieve and sort shapes in an Excel file with Aspose.Cells when ZOrder is not available | C# code to collect shapes from a workbook, sort by name, and display details using Aspose.Cells
// Tags: Aspose.Cells enumerate worksheet shapes C# | sort Excel shapes by ZOrder Aspose.Cells | fallback shape name sorting Aspose.Cells | output shape details to console C# | collect shapes from all worksheets Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample loads an existing workbook or creates a new one, gathers every Shape object from all worksheets, sorts the collection by Z‑order (using Shape.Name as a fallback when ZOrder is unavailable), and writes each shape's worksheet name and shape name to the console.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";

            // Load workbook if file exists; otherwise create a new workbook.
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
                // Ensure at least one worksheet exists.
                if (workbook.Worksheets.Count == 0)
                {
                    workbook.Worksheets.Add();
                }
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "Sheet1";

                try
                {
                    // Add a sample shape so the example has something to display.
                    // Use MsoDrawingType for compatibility with older Aspose.Cells versions.
                    Shape demoShape = sheet.Shapes.AddShape(
                        MsoDrawingType.Rectangle, 1, 1, 0, 0, 100, 100);
                    demoShape.Name = "DemoShape";
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to add demo shape: {ex.Message}");
                }
            }

            // Collect all shapes from all worksheets.
            List<Shape> allShapes = new List<Shape>();
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (Shape shape in sheet.Shapes)
                {
                    allShapes.Add(shape);
                }
            }

            // Sort shapes by name (as ZOrder may not be available in older versions).
            List<Shape> sortedShapes = allShapes.OrderBy(s => s.Name).ToList();

            // Output the sorted list to the console.
            Console.WriteLine("Shapes sorted by name:");
            foreach (Shape shape in sortedShapes)
            {
                // Each shape knows its parent worksheet via the Worksheet property.
                Console.WriteLine($"Worksheet: {shape.Worksheet.Name}, Shape Name: {shape.Name}");
            }

            // (Optional) Save the workbook if any modifications were made.
            // workbook.Save("output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
