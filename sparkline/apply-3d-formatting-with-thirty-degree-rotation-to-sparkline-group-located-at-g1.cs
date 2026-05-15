using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

class Apply3DToSparklineGroup
{
    static void Main()
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
        // G column = index 6 (0‑based), row 0
        CellArea location = new CellArea
        {
            StartRow = 0,
            EndRow = 0,
            StartColumn = 6,
            EndColumn = 6
        };

        // Add a line sparkline group with the data range and location
        int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line, "A1:D1", false, location);
        SparklineGroup sparkGroup = sheet.SparklineGroups[groupIndex];

        // (Optional) Add the sparkline explicitly – the Add method already creates it,
        // but this demonstrates the typical usage.
        sparkGroup.Sparklines.Add(sheet.Name + "!A1:D1", 0, 6);

        // To apply 3‑D formatting, we use a shape that occupies the same cell (G1).
        // The shape will be a simple rectangle; its ThreeDFormat allows rotation.
        // Position the shape roughly over cell G1 (column 6, row 0).
        // Parameters: upper left column, upper left row, lower right column, lower right row,
        // and the shape's width/height in pixels.
        Shape shape = sheet.Shapes.AddShape(
            MsoDrawingType.Rectangle, // shape type
            6,   // upper left column (G)
            0,   // upper left row
            6,   // lower right column (same column to keep it within one cell)
            0,   // lower right row
            100, // width in pixels
            20   // height in pixels
        );

        // Apply 3‑D formatting: set rotation around the Z‑axis to 30 degrees
        ThreeDFormat threeD = shape.ThreeDFormat;
        threeD.RotationZ = 30; // thirty‑degree rotation

        // (Optional) Add some visual cues so the 3‑D effect is noticeable
        threeD.ExtrusionHeight = 10;
        threeD.PresetCameraType = PresetCameraType.OrthographicFront;

        // Save the workbook
        workbook.Save("Sparkline3D_G1.xlsx");
    }
}