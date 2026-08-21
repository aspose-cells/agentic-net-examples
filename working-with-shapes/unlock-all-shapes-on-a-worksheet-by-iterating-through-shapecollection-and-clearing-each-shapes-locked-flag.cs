// Title: Unlock All Shapes in an Aspose.Cells Worksheet (C#) – Iterate ShapeCollection and Clear IsLocked
// Description: Shows how to create a workbook, add sample shapes, retrieve the worksheet's ShapeCollection, loop through each Shape, set its IsLocked property to false, and save the file as UnlockedShapes.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells unlock shapes C# | ShapeCollection iteration | IsLocked property false | unlock Excel shapes programmatically | Aspose.Cells shape protection | C# Aspose.Cells example | unlock worksheet shapes
// Common Searches: how to unlock all shapes in Aspose.Cells C# | iterate ShapeCollection to remove shape lock | set IsLocked false for Excel shapes using Aspose | C# code to unlock shapes in a workbook | Aspose.Cells unlock shape protection
// Developer Intent: Programmatically remove the lock from every shape on a worksheet.
// Use Cases: Prepare a template where end‑users can move or edit shapes freely after the workbook is generated. | Enable dynamic layout adjustments by unlocking shapes before repositioning them in a report. | Remove shape protection prior to exporting to formats that ignore locked‑shape settings.
// AI Prompts: Generate C# code with Aspose.Cells that unlocks all shapes in a worksheet and saves the workbook. | Explain step‑by‑step how to iterate through a ShapeCollection and set IsLocked = false for each shape. | Provide a concise example that creates a workbook, adds shapes, unlocks them, and writes UnlockedShapes.xlsx.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Shows how to create a workbook, add sample shapes, retrieve the worksheet's ShapeCollection, loop through each Shape, set its IsLocked property to false, and save the file as UnlockedShapes.xlsx using Aspose.Cells for .NET.
class UnlockShapesDemo
{
    static void Main()
    {
        // Create a new workbook (using the standard create rule)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Add some shapes to demonstrate unlocking (optional)
        sheet.Shapes.AddRectangle(2, 0, 2, 0, 100, 50);
        sheet.Shapes.AddOval(5, 0, 5, 0, 80, 80);

        // Get the ShapeCollection from the worksheet
        ShapeCollection shapes = sheet.Shapes;

        // Iterate through each shape and set its IsLocked property to false
        for (int i = 0; i < shapes.Count; i++)
        {
            Shape shape = shapes[i];
            shape.IsLocked = false; // Unlock the shape
        }

        // Save the workbook (using the standard save rule)
        workbook.Save("UnlockedShapes.xlsx", SaveFormat.Xlsx);
    }
}
