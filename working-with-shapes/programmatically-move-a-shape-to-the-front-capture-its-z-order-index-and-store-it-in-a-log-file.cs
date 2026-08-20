// Title: Bring a Shape to Front, Get Z‑Order Index, and Log It with Aspose.Cells for .NET
// Description: Shows how to add overlapping shapes in an Excel worksheet, move a shape to the front with ToFrontOrBack, read its ZOrderPosition, and append the index to a text file before saving the workbook.
// Keywords: Aspose.Cells | .NET | C# | shape Z-order | ToFrontOrBack | log Z-order to file | Excel shape layering | retrieve ZOrderPosition | move shape forward | Aspose.Cells example
// Common Searches: Aspose.Cells move shape to front C# | How to get Z-order of a shape in Aspose.Cells | Log shape Z-order index to text file .NET | ToFrontOrBack method Aspose.Cells example | Retrieve ZOrderPosition of overlapping shapes
// Developer Intent: The developer needs to bring a specific shape forward, read its Z‑order position, and record that value in a log file.
// Use Cases: Ensure a critical annotation appears above all other objects before exporting the workbook. | Audit and document the layering order of dynamically generated shapes in reports. | Debug overlapping shapes by writing their Z‑order indices to a log during workbook creation.
// AI Prompts: Provide C# code that moves a shape to the front and logs its ZOrderPosition using Aspose.Cells. | Explain how ToFrontOrBack works and how to retrieve the current Z-order of any shape in a worksheet. | Generate a script that records shape Z-order indices to a CSV file after reordering them.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShapeZOrderDemo
{
    // Shows how to add overlapping shapes in an Excel worksheet, move a shape to the front with ToFrontOrBack, read its ZOrderPosition, and append the index to a text file before saving the workbook.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add two overlapping shapes
            Shape shape1 = worksheet.Shapes.AddRectangle(10, 10, 100, 100, 0, 0);
            Shape shape2 = worksheet.Shapes.AddRectangle(50, 50, 100, 100, 0, 0);

            // Bring shape2 to the front (positive value moves forward)
            shape2.ToFrontOrBack(1);

            // Capture the Z-order position of shape2 after moving it
            int zOrderIndex = shape2.ZOrderPosition;

            // Log the Z-order index to a text file
            string logPath = "ShapeZOrderLog.txt";
            string logMessage = $"Shape2 Z-Order Position: {zOrderIndex}{Environment.NewLine}";
            File.AppendAllText(logPath, logMessage);

            // Save the workbook (optional, just to persist the changes)
            workbook.Save("ShapeZOrderDemo.xlsx");
        }
    }
}
