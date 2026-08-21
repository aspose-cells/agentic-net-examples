// Title: Group Multiple Shapes into a Named GroupShape (IconSet) with Aspose.Cells for .NET (C#)
// Description: C# example that loads or creates an Excel workbook, locates shapes named Icon1, Icon2, and Icon3, groups them into a single GroupShape called IconSet using Aspose.Cells, and saves the workbook. Includes fallback logic to add placeholder shapes when they are missing.
// Keywords: Aspose.Cells | C# | .NET | group shapes | GroupShape | Excel shape grouping | IconSet | worksheet.Shapes.Group | Aspose.Cells example | GitHub
// Common Searches: Aspose.Cells group shapes C# | How to create a GroupShape in Excel using Aspose.Cells | C# group Icon1 Icon2 Icon3 into one group | Aspose.Cells shape grouping tutorial | GroupShape IconSet example | Combine multiple shapes Aspose.Cells .NET | GitHub Aspose.Cells shape grouping sample
// Developer Intent: Create a GroupShape named IconSet that contains the existing shapes Icon1, Icon2, and Icon3 so they can be manipulated together.
// Use Cases: Move, resize, or rotate all three icons as a single unit | Toggle visibility or lock the entire icon set with one property | Apply formatting or data binding to the grouped icons simultaneously | Copy or export the grouped icons to another worksheet or workbook
// AI Prompts: Write C# code using Aspose.Cells to find shapes Icon1, Icon2, Icon3, group them into a GroupShape named IconSet, and then rotate the group by 45 degrees. | Provide a robust method that checks for required shapes, creates missing placeholders, groups them, and logs warnings without throwing exceptions. | Show how to ungroup a GroupShape in Aspose.Cells and iterate over its child shapes to modify each individually. | Create a GitHub‑compatible snippet that demonstrates shape grouping with clear comments for developers.

using System;
using System.IO;
using System.Linq;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShapeGrouping
{
    // C# example that loads or creates an Excel workbook, locates shapes named Icon1, Icon2, and Icon3, groups them into a single GroupShape called IconSet using Aspose.Cells, and saves the workbook. Includes fallback logic to add placeholder shapes when they are missing.
    class Program
    {
        static void Main()
        {
            try
            {
                const string inputPath = "InputWorkbook.xlsx";
                const string outputPath = "OutputWorkbook.xlsx";

                // Load existing workbook or create a new one if the file does not exist
                Workbook workbook;
                if (File.Exists(inputPath))
                {
                    workbook = new Workbook(inputPath);
                }
                else
                {
                    workbook = new Workbook();
                    // Add three placeholder shapes so the grouping logic can run
                    Worksheet ws = workbook.Worksheets[0];
                    AddPlaceholderShape(ws, "Icon1", 1, 1);
                    AddPlaceholderShape(ws, "Icon2", 3, 1);
                    AddPlaceholderShape(ws, "Icon3", 5, 1);
                }

                Worksheet worksheet = workbook.Worksheets[0];

                // Retrieve the shapes named Icon1, Icon2, and Icon3
                Shape[] icons = worksheet.Shapes
                    .Cast<Shape>()
                    .Where(s => s.Name == "Icon1" || s.Name == "Icon2" || s.Name == "Icon3")
                    .ToArray();

                // Ensure all three shapes were found before grouping
                if (icons.Length == 3)
                {
                    // Group the three shapes
                    GroupShape iconSet = worksheet.Shapes.Group(icons);
                    iconSet.Name = "IconSet";
                }
                else
                {
                    Console.WriteLine("Warning: One or more required shapes (Icon1, Icon2, Icon3) were not found. No grouping performed.");
                }

                // Save the workbook with the new group (or unchanged if grouping was skipped)
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Helper method to add a simple rectangle shape as a placeholder
        private static void AddPlaceholderShape(Worksheet ws, string name, int row, int column)
        {
            // Add a rectangle shape and obtain the shape object directly
            Shape shape = ws.Shapes.AddShape(MsoDrawingType.Rectangle, row, column, 0, 0, 100, 50);
            shape.Name = name;
            shape.Placement = PlacementType.FreeFloating;
        }
    }
}
