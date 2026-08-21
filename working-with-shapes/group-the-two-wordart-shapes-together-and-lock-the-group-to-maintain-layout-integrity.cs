// Title: Group and lock WordArt shapes in Aspose.Cells for .NET (C#)
// Description: Demonstrates how to add two WordArt‑style rectangles to a worksheet, combine them with ShapeCollection.Group, lock the resulting GroupShape (including aspect‑ratio lock), and save the workbook as GroupedWordArt.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# | group shapes Aspose.Cells | lock shape group | WordArt Excel C# | ShapeCollection.Group | IsLocked property | IsAspectRatioLocked | Excel shape protection | C# workbook example
// Common Searches: Aspose.Cells group multiple shapes C# | lock a shape group in Excel with Aspose | prevent shape movement after grouping Aspose.Cells | set aspect ratio lock for grouped shapes .NET | WordArt grouping example Aspose.Cells
// Developer Intent: Combine two WordArt objects into a single group and make the group immutable so its position and size cannot be altered.
// Use Cases: Create a title‑subtitle pair on a report sheet and lock them together to stay aligned when the sheet is protected. | Assemble a multi‑part logo from separate shapes, group it, and prevent accidental repositioning or resizing. | Design a dashboard banner with decorative elements, lock the group’s aspect ratio, and distribute the workbook without layout changes.
// AI Prompts: Show C# code to group three shapes and lock the group with Aspose.Cells. | Explain how to disable resizing of a locked shape group on a protected worksheet using Aspose.Cells for .NET. | Provide steps to ungroup shapes, edit their text, and regroup them with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsGroupWordArtDemo
{
    // Demonstrates how to add two WordArt‑style rectangles to a worksheet, combine them with ShapeCollection.Group, lock the resulting GroupShape (including aspect‑ratio lock), and save the workbook as GroupedWordArt.xlsx using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Get the shapes collection of the worksheet
            ShapeCollection shapes = sheet.Shapes;

            // Add two sample WordArt-like shapes (using rectangles for demonstration)
            // Parameters: upperLeftRow, upperLeftColumn, top, left, height, width
            Shape wordArt1 = shapes.AddRectangle(2, 0, 2, 0, 60, 200);
            Shape wordArt2 = shapes.AddRectangle(6, 0, 2, 0, 60, 200);

            // Optionally set some text to mimic WordArt
            wordArt1.Text = "Hello";
            wordArt2.Text = "World";

            // Group the two shapes together
            GroupShape group = shapes.Group(new Shape[] { wordArt1, wordArt2 });

            // Lock the group to preserve its layout (prevents moving/resizing when sheet is protected)
            group.IsLocked = true;

            // Optionally lock aspect ratio as an extra safeguard
            group.IsAspectRatioLocked = true;

            // Save the workbook
            workbook.Save("GroupedWordArt.xlsx");
        }
    }
}
