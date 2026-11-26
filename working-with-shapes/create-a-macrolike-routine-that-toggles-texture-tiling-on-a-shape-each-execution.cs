// Title: C# routine to toggle texture fill tiling on a rectangle shape in an Excel workbook using Aspose.Cells
// Common Searches: how to programmatically toggle texture fill tiling for a shape in Aspose.Cells C# | Aspose.Cells C# change IsTiling property of a shape's texture fill | toggle rectangle shape texture tiling each run with Aspose.Cells .NET | C# example to find first rectangle shape and switch texture tiling in an Excel file | Aspose.Cells macro-like code to enable or disable texture tiling on shapes
// Tags: Aspose.Cells texture fill IsTiling example | C# add rectangle shape with texture fill | Aspose.Cells locate first rectangle shape | Excel workbook shape texture tiling C# | Aspose.Cells shape fill type texture

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace TextureTilingToggle
{
    // The program loads an existing workbook or creates a new one, locates the first rectangle shape (or adds one), sets its fill type to texture, flips the TextureFill.IsTiling flag, and saves the workbook, providing a macro‑like routine to toggle texture tiling on each execution.
    class Program
    {
        // Path to the Excel file that will be used for the demo.
        // Change this path as needed.
        private const string WorkbookPath = "TextureTilingDemo.xlsx";

        static void Main(string[] args)
        {
            try
            {
                Workbook workbook;
                Worksheet worksheet;
                Shape shape;

                // Load existing workbook or create a new one.
                if (File.Exists(WorkbookPath))
                {
                    Console.WriteLine($"Loading existing workbook: {WorkbookPath}");
                    workbook = new Workbook(WorkbookPath);
                    worksheet = workbook.Worksheets[0];

                    // Try to find the first rectangle shape in the worksheet.
                    shape = FindFirstRectangleShape(worksheet);
                    if (shape == null)
                    {
                        Console.WriteLine("No rectangle shape found in the existing workbook. Adding a new rectangle.");
                        shape = AddRectangleShape(worksheet);
                    }
                }
                else
                {
                    Console.WriteLine("Workbook not found. Creating a new workbook with a rectangle shape.");
                    workbook = new Workbook();
                    worksheet = workbook.Worksheets[0];
                    shape = AddRectangleShape(worksheet);
                }

                // Ensure the shape uses texture fill.
                shape.Fill.FillType = FillType.Texture;
                TextureFill textureFill = shape.Fill.TextureFill;

                // Toggle the IsTiling property.
                bool currentTiling = textureFill.IsTiling;
                textureFill.IsTiling = !currentTiling;

                Console.WriteLine($"Texture tiling was {(currentTiling ? "enabled" : "disabled")}. It is now {(textureFill.IsTiling ? "enabled" : "disabled")}.");

                // Save the workbook.
                workbook.Save(WorkbookPath);
                Console.WriteLine($"Workbook saved to: {WorkbookPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while toggling texture tiling:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }

        private static Shape AddRectangleShape(Worksheet worksheet)
        {
            // Add a rectangle shape directly; the method returns a RectangleShape instance.
            RectangleShape rect = worksheet.Shapes.AddRectangle(2, 2, 0, 0, 150, 100);

            // Set a built‑in texture so that tiling effect is visible.
            rect.Fill.FillType = FillType.Texture;
            TextureFill tf = rect.Fill.TextureFill;
            tf.Type = TextureType.BlueTissuePaper; // Any built‑in texture works.
            tf.IsTiling = true; // Initial state; will be toggled on first run.

            Console.WriteLine("Added rectangle shape with initial texture tiling set to true.");
            return rect;
        }

        private static Shape FindFirstRectangleShape(Worksheet worksheet)
        {
            foreach (Shape s in worksheet.Shapes)
            {
                // Rectangle shapes are of type ShapeType.Rectangle.
                if (s.AutoShapeType == AutoShapeType.Rectangle)
                {
                    return s;
                }
            }
            return null;
        }
    }
}
