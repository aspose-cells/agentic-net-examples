// Title: C# Aspose.Cells: List and Alphabetically Sort Worksheet Shapes by Z‑Order Index
// Description: Loads a workbook with Aspose.Cells for .NET, extracts each shape's Name and ZOrderPosition from the first worksheet, orders the entries alphabetically (case‑insensitive), prints a tab‑delimited report, and saves the file unchanged.
// Keywords: Aspose.Cells shape Z order | C# list Excel shapes | Aspose.Cells ZOrderPosition | sort shapes by name | Excel shape report .NET | console shape audit | shape layering Aspose | C# Aspose.Cells example
// Common Searches: Aspose.Cells get shape Z‑order C# | list Excel shapes alphabetically Aspose | C# code to report shape order in workbook | how to sort shapes by name using Aspose.Cells | retrieve shape names and ZOrderPosition .NET
// Developer Intent: Extract each shape’s name and Z‑order index from a worksheet and output the list sorted by name.
// Use Cases: Audit the layering sequence of charts, images, and text boxes before publishing a workbook. | Generate documentation that maps shape names to their Z‑order for design reviews. | Validate naming conventions and proper Z‑order placement during automated Excel generation.
// AI Prompts: Create a reusable method that returns a List<(string Name, int ZOrder)> of all shapes on a worksheet, sorted alphabetically. | Modify the example to write the sorted shape report to a CSV or JSON file instead of the console. | Add robust error handling that logs unnamed shapes and missing Z‑order values while still producing the sorted output.

using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace ShapeZOrderReporter
{
    // Loads a workbook with Aspose.Cells for .NET, extracts each shape's Name and ZOrderPosition from the first worksheet, orders the entries alphabetically (case‑insensitive), prints a tab‑delimited report, and saves the file unchanged.
    class Program
    {
        static void Main(string[] args)
        {
            // Load an existing workbook (replace with actual path)
            string inputPath = "input.xlsx";
            Workbook workbook = new Workbook(inputPath);

            // Choose the worksheet to analyze (first worksheet in this example)
            Worksheet worksheet = workbook.Worksheets[0];

            // Collect shape name and Z‑order position
            List<(string Name, int ZOrder)> shapeInfo = new List<(string, int)>();

            foreach (Shape shape in worksheet.Shapes)
            {
                // Ensure a name is available; use empty string if null
                string name = shape.Name ?? string.Empty;
                shapeInfo.Add((name, shape.ZOrderPosition));
            }

            // Sort the collected information alphabetically by shape name
            var sortedShapeInfo = shapeInfo.OrderBy(s => s.Name, StringComparer.OrdinalIgnoreCase);

            // Output the report
            Console.WriteLine("Shape Name\tZ‑Order Position");
            Console.WriteLine("-----------------------------------");
            foreach (var info in sortedShapeInfo)
            {
                Console.WriteLine($"{info.Name}\t{info.ZOrder}");
            }

            // Save the workbook (unchanged) – replace with desired output path
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
        }
    }
}
