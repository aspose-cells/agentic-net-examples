// Title: Unlock Every Shape in an Excel Worksheet with Aspose.Cells (C#)
// Description: This C# sample loads a workbook, accesses a worksheet's ShapeCollection, sets each shape's IsLocked property to false, and saves the file, removing shape protection from the sheet.
// Keywords: Aspose.Cells | C# unlock shapes | Excel shape IsLocked | ShapeCollection iteration | remove shape lock | Aspose.Cells API | unlock worksheet shapes
// Common Searches: how to unlock shapes in Excel using Aspose.Cells | C# code to clear shape lock flag Aspose.Cells | iterate ShapeCollection to change IsLocked property | unlock all charts and images in an Excel file programmatically | Aspose.Cells batch unlock worksheet shapes
// Developer Intent: Clear the IsLocked flag for every shape on a worksheet.
// Use Cases: Prepare template workbooks so end users can edit charts, pictures, and text boxes without restrictions. | Batch‑process multiple Excel files to ensure no shape is locked before applying bulk formatting. | Enable automated data imports that need to move or resize shapes programmatically. | Remove shape protection prior to running macro‑free visual updates.
// AI Prompts: Write a C# method that receives a Worksheet object and unlocks all its shapes using Aspose.Cells. | Provide code to unlock shapes on every sheet of a workbook, handling empty ShapeCollections gracefully. | Create a reusable utility that toggles the IsLocked property of shapes based on a boolean parameter and returns the modified workbook. | Generate a PowerShell script that uses Aspose.Cells to unlock shapes in all Excel files within a directory.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// This C# sample loads a workbook, accesses a worksheet's ShapeCollection, sets each shape's IsLocked property to false, and saves the file, removing shape protection from the sheet.
class UnlockAllShapes
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (or any specific worksheet)
        Worksheet worksheet = workbook.Worksheets[0];

        // Get the collection of shapes on the worksheet
        ShapeCollection shapes = worksheet.Shapes;

        // Iterate through each shape and unlock it
        for (int i = 0; i < shapes.Count; i++)
        {
            Shape shape = shapes[i];
            shape.IsLocked = false; // Unlock the shape
        }

        // Save the modified workbook
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}
