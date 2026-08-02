// Title: List and Sort Worksheet Shape Z‑Order Indices with Aspose.Cells for .NET
// Description: A C# utility that creates a workbook, adds named shapes, adjusts their Z‑order, extracts each shape's Name and ZOrderPosition, sorts the entries alphabetically, prints a tabular report, and saves the file as ShapeZOrderReport.xlsx.
// Keywords: Aspose.Cells shape Z-order | C# list worksheet shapes | retrieve Shape.ZOrderPosition | sort Excel shapes by name | shape hierarchy report .NET | enumerate shapes in workbook | Aspose.Cells shape reporting utility
// Common Searches: how to get shape Z-order in Aspose.Cells | sort Excel shapes alphabetically using Aspose.Cells | list shape names and Z-order positions .NET | Aspose.Cells shape layering report example | retrieve and display shape ZOrderPosition C#
// Developer Intent: Generate a report that shows each worksheet shape’s Z‑order index, ordered alphabetically by shape name.
// Use Cases: Create a printable audit of shape layering before publishing the workbook. | Validate shape order programmatically to ensure correct rendering when converting to PDF. | Automate detection of unintended Z‑order changes in generated spreadsheets.
// AI Prompts: Write a method that returns a DataTable with shape names and Z‑order values sorted alphabetically. | Modify the example to export the sorted shape list to a CSV file instead of the console. | Add robust error handling for worksheets without shapes and log the sorted report to a file.

using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace ShapeZOrderReport
{
    // A C# utility that creates a workbook, adds named shapes, adjusts their Z‑order, extracts each shape's Name and ZOrderPosition, sorts the entries alphabetically, prints a tabular report, and saves the file as ShapeZOrderReport.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle create)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add sample shapes with distinct names
            Shape rect = sheet.Shapes.AddRectangle(2, 0, 2, 0, 80, 120);
            rect.Name = "RectangleA";

            Shape oval = sheet.Shapes.AddOval(5, 0, 5, 0, 80, 120);
            oval.Name = "OvalB";

            Shape line = sheet.Shapes.AddLine(8, 0, 8, 0, 80, 120);
            line.Name = "LineC";

            // Change Z-order to demonstrate different positions
            // Bring "OvalB" to front (higher Z-order)
            oval.ToFrontOrBack(1);
            // Send "LineC" to back (lower Z-order)
            line.ToFrontOrBack(-1);

            // Retrieve all shapes from the worksheet
            ShapeCollection shapes = sheet.Shapes;

            // Build a list of shape info (name and Z-order)
            List<(string Name, int ZOrder)> shapeInfo = new List<(string, int)>();
            foreach (Shape shape in shapes)
            {
                // Use Shape.Name and Shape.ZOrderPosition properties
                shapeInfo.Add((shape.Name, shape.ZOrderPosition));
            }

            // Sort the list alphabetically by shape name
            var sortedInfo = shapeInfo.OrderBy(s => s.Name, StringComparer.OrdinalIgnoreCase);

            // Output the report
            Console.WriteLine("Shape Name\tZ-Order Position");
            Console.WriteLine("--------------------------------");
            foreach (var info in sortedInfo)
            {
                Console.WriteLine($"{info.Name}\t{info.ZOrder}");
            }

            // Save the workbook (lifecycle save)
            workbook.Save("ShapeZOrderReport.xlsx");
        }
    }
}
