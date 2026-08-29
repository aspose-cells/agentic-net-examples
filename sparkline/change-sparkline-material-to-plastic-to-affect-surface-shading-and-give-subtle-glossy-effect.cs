// Title: Set a rectangle shape’s 3‑D material to Plastic while adding a line sparkline in Aspose.Cells for .NET (C#)
// AI Prompts: Generate a C# Aspose.Cells sample that builds a workbook, creates a line sparkline from a data range, adds a rectangle shape, and applies the Plastic preset material to the shape’s 3‑D format. | Show how to configure extrusion height, X‑axis rotation, Y‑axis rotation, and lighting for a shape’s 3‑D format to produce a subtle glossy appearance alongside a sparkline.
// Common Searches: how to apply Plastic material to a shape using Aspose.Cells in C# | C# Aspose.Cells adjust shape extrusion height and rotation | example of adding a line sparkline and a 3‑D rectangle shape in an Excel workbook with Aspose.Cells | Aspose.Cells tutorial for creating glossy Plastic material effect on shapes | Aspose.Cells C# set shape ThreeDFormat material Plastic
// Tags: Aspose.Cells set shape 3d material plastic | C# create line sparkline Aspose.Cells | ThreeDFormat extrusion rotation lighting example | rectangle shape glossy effect Excel Aspose.Cells | presetmaterialtype plastic usage Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// The example creates a new workbook, adds a line sparkline based on cells A1:D1, inserts a rectangle shape, and sets the shape's ThreeDFormat.Material to the Plastic preset. It also demonstrates configuring extrusion height, X/Y rotation, and lighting to achieve a subtle glossy effect, then saves the file as SparklineWithPlasticMaterialDemo.xlsx.
class SparklineMaterialDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the sparkline
        sheet.Cells["A1"].PutValue(5);
        sheet.Cells["B1"].PutValue(2);
        sheet.Cells["C1"].PutValue(1);
        sheet.Cells["D1"].PutValue(3);

        // Define the location where the sparkline will be placed
        CellArea sparklineLocation = new CellArea
        {
            StartRow = 0,
            EndRow = 0,
            StartColumn = 4,
            EndColumn = 4
        };

        // Add a line sparkline group
        int sparklineGroupIndex = sheet.SparklineGroups.Add(
            SparklineType.Line,
            "A1:D1",
            false,
            sparklineLocation);

        SparklineGroup sparklineGroup = sheet.SparklineGroups[sparklineGroupIndex];
        sparklineGroup.Sparklines.Add(sheet.Name + "!A1:D1", 0, 4);

        // OPTIONAL: Customize sparkline appearance
        CellsColor seriesColor = workbook.CreateCellsColor();
        seriesColor.Color = Color.Orange;
        sparklineGroup.SeriesColor = seriesColor;
        sparklineGroup.ShowMarkers = true;

        // Add a shape to demonstrate 3‑D material effect
        // The shape itself is not a sparkline, but we can illustrate the Plastic material here
        Shape shape = sheet.Shapes.AddShape(
            MsoDrawingType.Rectangle, // shape type
            5,   // upper left row
            5,   // upper left column
            0,   // upper left offset in pixels
            0,   // upper left offset in pixels
            150, // width in pixels
            50   // height in pixels
        );

        // Set the 3‑D material of the shape to Plastic for a subtle glossy effect
        shape.ThreeDFormat.Material = PresetMaterialType.Plastic;

        // Optionally configure additional 3‑D properties to better see the effect
        shape.ThreeDFormat.ExtrusionHeight = 10;
        shape.ThreeDFormat.RotationX = 15;
        shape.ThreeDFormat.RotationY = 30;
        shape.ThreeDFormat.Lighting = LightRigType.BrightRoom;

        // Save the workbook
        workbook.Save("SparklineWithPlasticMaterialDemo.xlsx");
    }
}
