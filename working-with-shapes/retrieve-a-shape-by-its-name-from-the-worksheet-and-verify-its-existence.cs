// Title: How to retrieve a named shape and verify its existence in an Excel worksheet using Aspose.Cells for .NET
// Common Searches: asp.net retrieve shape by name from worksheet using Aspose.Cells | c# check if an auto shape exists in an Excel file with Aspose.Cells | find rectangle shape MyRectangleShape in an Aspose.Cells workbook | case‑insensitive shape lookup in Aspose.Cells C# example | load workbook and get specific shape by name Aspose.Cells
// Tags: shape name lookup Aspose.Cells .NET | verify worksheet shape presence C# | Aspose.Cells auto shape retrieval example | excel shape existence verification Aspose | case‑insensitive shape matching Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using System.Drawing;

namespace RetrieveShapeByNameDemo
{
    // The sample loads or creates an Excel workbook, ensures a rectangle auto‑shape named "MyRectangleShape" is present on the first worksheet, then iterates the worksheet's Shapes collection to locate the shape by its Name property (case‑insensitive), outputs its type, position and size, and finally saves the workbook.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the workbook file. Change as needed.
            string workbookPath = "ShapeDemo.xlsx";

            try
            {
                // Load existing workbook if it exists; otherwise create a new one.
                Workbook workbook = File.Exists(workbookPath) ? new Workbook(workbookPath) : new Workbook();
                Console.WriteLine(File.Exists(workbookPath)
                    ? $"Loaded workbook from \"{workbookPath}\"."
                    : "Created a new workbook.");

                // Use the first worksheet.
                Worksheet worksheet = workbook.Worksheets[0];
                worksheet.Name = "DemoSheet";

                // ------------------------------------------------------------
                // Ensure a shape with a known name exists.
                // ------------------------------------------------------------
                const string targetShapeName = "MyRectangleShape";

                // Try to retrieve the shape first (in case the workbook already contains it).
                Shape existingShape = FindShapeByName(worksheet, targetShapeName);
                if (existingShape != null)
                {
                    Console.WriteLine($"Shape \"{targetShapeName}\" already exists in the worksheet.");
                }
                else
                {
                    // Add a rectangle auto shape and assign the target name.
                    // Parameters: AutoShapeType, upper left row, upper left column,
                    // top offset (pixels), left offset (pixels), height (pixels), width (pixels)
                    Shape newShape = worksheet.Shapes.AddAutoShape(
                        AutoShapeType.Rectangle,
                        2,          // top row index (zero‑based)
                        2,          // left column index (zero‑based)
                        0,          // top offset in pixels
                        0,          // left offset in pixels
                        100,        // height in pixels
                        200);       // width in pixels

                    newShape.Name = targetShapeName;
                    newShape.AlternativeText = "Demo rectangle shape";
                    newShape.Fill.SolidFill.Color = Color.LightBlue;
                    newShape.Line.Weight = 2;

                    Console.WriteLine($"Added shape \"{targetShapeName}\" to the worksheet.");
                }

                // ------------------------------------------------------------
                // Retrieve the shape by name and verify its existence.
                // ------------------------------------------------------------
                Shape retrievedShape = FindShapeByName(worksheet, targetShapeName);

                if (retrievedShape != null)
                {
                    Console.WriteLine($"Successfully retrieved shape \"{targetShapeName}\".");
                    Console.WriteLine($" - Shape Type: {retrievedShape.Type}");
                    Console.WriteLine($" - Position: Row {retrievedShape.UpperLeftRow}, Column {retrievedShape.UpperLeftColumn}");
                    Console.WriteLine($" - Size: Height={retrievedShape.Height} px, Width={retrievedShape.Width} px");
                }
                else
                {
                    Console.WriteLine($"Shape \"{targetShapeName}\" was NOT found in the worksheet.");
                }

                // Save the workbook (overwrites the original file if it existed).
                workbook.Save(workbookPath);
                Console.WriteLine($"Workbook saved to \"{workbookPath}\".");
            }
            catch (Exception ex)
            {
                // Output full exception details for debugging.
                Console.WriteLine("An error occurred while processing the workbook:");
                Console.WriteLine(ex.ToString());
            }
        }

        private static Shape FindShapeByName(Worksheet worksheet, string shapeName)
        {
            foreach (Shape shape in worksheet.Shapes)
            {
                if (shape.Name.Equals(shapeName, StringComparison.OrdinalIgnoreCase))
                {
                    return shape;
                }
            }
            return null;
        }
    }
}
