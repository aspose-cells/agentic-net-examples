// Title: Ungroup Shapes in Aspose.Cells for .NET (C#) While Preserving Fill and Line Formatting
// Description: Demonstrates how to group two rectangle shapes, ungroup them with GroupShape.Ungroup, and keep each shape's original fill and line colors. The example shows the IsInGroup flag before and after ungrouping and saves the workbook as an Excel file.
// Keywords: Aspose.Cells ungroup shapes | C# group shape ungroup | preserve shape formatting Aspose.Cells | GroupShape.Ungroup example | IsInGroup property | .NET Excel shape formatting
// Common Searches: how to ungroup shapes in Aspose.Cells C# | retain fill color after ungrouping Aspose.Cells | Aspose.Cells GroupShape.Ungroup usage | check IsInGroup flag Aspose.Cells | C# example ungroup Excel shapes
// Developer Intent: Separate previously grouped shapes so each retains its original fill and line settings.
// Use Cases: Create a workbook, group multiple shapes, then ungroup to edit them individually without losing colors. | Verify that the IsInGroup property changes from true to false after calling Ungroup. | Generate an Excel file where each shape appears exactly as designed after ungrouping.
// AI Prompts: Write C# code using Aspose.Cells to ungroup a GroupShape while keeping each shape's FillFormat and LineFormat. | Explain the effect of GroupShape.Ungroup on the IsInGroup property of member shapes. | Provide a step‑by‑step example that groups, ungroups, and saves shapes in an Excel workbook without altering their appearance.

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsUngroupShapesDemo
{
    // Demonstrates how to group two rectangle shapes, ungroup them with GroupShape.Ungroup, and keep each shape's original fill and line colors. The example shows the IsInGroup flag before and after ungrouping and saves the workbook as an Excel file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add two rectangle shapes with distinct formatting
                Shape rect1 = worksheet.Shapes.AddRectangle(2, 0, 2, 0, 80, 40);
                rect1.FillFormat.ForeColor = Color.LightBlue;   // Set fill color
                rect1.LineFormat.ForeColor = Color.DarkBlue;    // Set line color
                rect1.Text = "Rect 1";

                Shape rect2 = worksheet.Shapes.AddRectangle(6, 0, 2, 0, 80, 40);
                rect2.FillFormat.ForeColor = Color.LightGreen;  // Set fill color
                rect2.LineFormat.ForeColor = Color.DarkGreen;   // Set line color
                rect2.Text = "Rect 2";

                // Group the two shapes
                GroupShape group = worksheet.Shapes.Group(new Shape[] { rect1, rect2 });

                // At this point both shapes are inside the group (IsInGroup == true)
                Console.WriteLine($"Before ungroup - Rect1 IsInGroup: {rect1.IsInGroup}");
                Console.WriteLine($"Before ungroup - Rect2 IsInGroup: {rect2.IsInGroup}");

                // Ungroup the shapes; formatting is preserved automatically
                group.Ungroup();

                // After ungrouping each shape should retain its original formatting
                Console.WriteLine($"After ungroup - Rect1 IsInGroup: {rect1.IsInGroup}");
                Console.WriteLine($"After ungroup - Rect2 IsInGroup: {rect2.IsInGroup}");
                Console.WriteLine($"Rect1 Fill Color: {rect1.FillFormat.ForeColor}");
                Console.WriteLine($"Rect2 Fill Color: {rect2.FillFormat.ForeColor}");

                // Save the workbook
                string outputPath = "UngroupedShapesDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
