// Title: C# – Apply a 30° 3‑D Rotation to a Sparkline Group in Cell G1 using Aspose.Cells
// Description: Creates a workbook, adds a line sparkline from A1:D1 to G1, overlays a semi‑transparent rectangle shape, sets ThreeDFormat.RotationZ to 30°, and saves the file as SparklineWith3DRotation.xlsx.
// Keywords: Aspose.Cells | C# | sparkline | 3D rotation | ThreeDFormat | RotationZ | overlay shape | transparent rectangle | Excel sparkline styling | G1 cell
// Common Searches: Aspose.Cells rotate sparkline 30 degrees C# | how to add 3D format to sparkline Aspose.Cells | overlay shape on sparkline cell Aspose.Cells | transparent shape over sparkline Excel C# | apply ThreeDFormat.RotationZ to sparkline group
// Developer Intent: Add a line sparkline to G1 and give it a 30° Z‑axis rotation by overlaying a semi‑transparent rectangle shape.
// Use Cases: Tilt a sparkline on a financial dashboard to create a 3‑D visual effect. | Combine sparkline data with shape styling for richer Excel reports. | Maintain sparkline readability while adding depth through a transparent overlay.
// AI Prompts: Generate C# code that inserts a line sparkline in cell G1 and rotates it 30 degrees using Aspose.Cells ThreeDFormat. | Show how to overlay a semi‑transparent rectangle on a sparkline cell and configure its RotationZ property. | Explain how to size and set transparency of a shape so the underlying sparkline stays visible after applying 3‑D rotation.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsSparkline3DExample
{
    // Creates a workbook, adds a line sparkline from A1:D1 to G1, overlays a semi‑transparent rectangle shape, sets ThreeDFormat.RotationZ to 30°, and saves the file as SparklineWith3DRotation.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the sparkline (A1:D1)
                sheet.Cells["A1"].PutValue(5);
                sheet.Cells["B1"].PutValue(3);
                sheet.Cells["C1"].PutValue(8);
                sheet.Cells["D1"].PutValue(2);

                // Define the location range for the sparkline group (cell G1)
                CellArea sparklineArea = new CellArea
                {
                    StartRow = 0,   // Row 1 (zero‑based)
                    EndRow = 0,
                    StartColumn = 6, // Column G (zero‑based, A=0)
                    EndColumn = 6
                };

                // Add a line sparkline group with the data range and location
                int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line, "A1:D1", false, sparklineArea);
                SparklineGroup sparklineGroup = sheet.SparklineGroups[groupIndex];

                // To apply 3‑D formatting, overlay a shape on the same cell and set its ThreeDFormat
                // The shape will visually represent the 30‑degree rotation.
                // Add a rectangle shape that covers the G1 cell
                Shape shape = sheet.Shapes.AddShape(
                    MsoDrawingType.Rectangle,
                    0,          // upper left row
                    6,          // upper left column (G)
                    0,          // upper left row offset (points)
                    0,          // upper left column offset (points)
                    100,        // width in points
                    20          // height in points
                );

                // Apply 3‑D rotation of 30 degrees around the Z‑axis
                shape.ThreeDFormat.RotationZ = 30;

                // Make the shape semi‑transparent so the sparkline remains visible
                shape.Fill.Transparency = 0.8; // 80% transparent

                // Save the workbook
                workbook.Save("SparklineWith3DRotation.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
