// Title: Aspose.Cells .NET: Log SmartArt Shape Type, Position, Size, Visibility and Z‑Order in Excel
// Description: C# sample that opens an Excel workbook, scans every worksheet, detects SmartArt shapes, and writes each shape's name, .NET type, left/top coordinates, width/height (pixels), hidden flag and Z‑order to the console.
// Keywords: Aspose.Cells SmartArt shape enumeration | C# get SmartArt coordinates | log SmartArt dimensions Aspose | retrieve SmartArt visibility .NET | Excel SmartArt Z‑order Aspose.Cells | Aspose.Cells shape properties | SmartArt shape diagnostics C#
// Common Searches: how to list SmartArt shapes in an Excel file using Aspose.Cells | Aspose.Cells get SmartArt position and size | enumerate SmartArt objects in .NET workbook | retrieve SmartArt Z‑order with Aspose.Cells | log hidden status of SmartArt shapes in C#
// Developer Intent: Extract and output the name, type, coordinates, dimensions, hidden state and Z‑order of every SmartArt shape contained in an Excel workbook.
// Use Cases: Create a layout audit report to verify SmartArt placement before publishing a spreadsheet. | Automate quality checks that flag SmartArt objects outside expected size or position ranges. | Debug visual stacking issues by reviewing each shape's Z‑order and visibility flag.
// AI Prompts: Generate C# code that writes the logged SmartArt details to a CSV file using Aspose.Cells. | Show how to offset each SmartArt shape by a given X/Y value after logging its original coordinates. | Explain how to filter SmartArt shapes by specific diagram types (e.g., Process, Cycle) with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace SmartArtInfoLogger
{
    // C# sample that opens an Excel workbook, scans every worksheet, detects SmartArt shapes, and writes each shape's name, .NET type, left/top coordinates, width/height (pixels), hidden flag and Z‑order to the console.
    class Program
    {
        static void Main(string[] args)
        {
            // Load an existing workbook (replace with your actual file path)
            string inputPath = "input.xlsx";
            Workbook workbook = new Workbook(inputPath);

            // Iterate through each worksheet
            foreach (Worksheet worksheet in workbook.Worksheets)
            {
                // Iterate through all shapes in the worksheet
                foreach (Shape shape in worksheet.Shapes)
                {
                    // Check if the shape is a SmartArt shape
                    if (shape.IsSmartArt && shape is SmartArtShape smartArtShape)
                    {
                        // Log basic identification
                        Console.WriteLine("=== SmartArt Shape Detected ===");
                        Console.WriteLine($"Name          : {shape.Name}");
                        Console.WriteLine($"Shape Type    : {smartArtShape.GetType().Name}");

                        // Log position (in pixels)
                        Console.WriteLine($"Left (X)      : {shape.Left}");
                        Console.WriteLine($"Top (Y)       : {shape.Top}");

                        // Log size (in pixels)
                        Console.WriteLine($"Width         : {shape.Width}");
                        Console.WriteLine($"Height        : {shape.Height}");

                        // Additional optional details
                        Console.WriteLine($"IsHidden      : {shape.IsHidden}");
                        Console.WriteLine($"Z-Order       : {shape.ZOrderPosition}");
                        Console.WriteLine();
                    }
                }
            }

            // Optionally save the workbook if any modifications were made
            // workbook.Save("output.xlsx");
        }
    }
}
