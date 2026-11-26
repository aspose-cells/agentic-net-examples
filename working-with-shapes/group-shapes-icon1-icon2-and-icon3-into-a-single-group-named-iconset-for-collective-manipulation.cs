// Title: How to group multiple rectangle shapes into a named GroupShape (IconSet) with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# group rectangle shapes into a single shape | Create a named GroupShape called IconSet using Aspose.Cells for .NET | Add several shapes and group them in an Excel workbook with Aspose.Cells | Set placement and line formatting for a GroupShape in Aspose.Cells C# | Save an Excel file that contains grouped shapes using Aspose.Cells
// Tags: group shapes using Aspose.Cells ShapeCollection.Group | named GroupShape creation Aspose.Cells C# | rectangle shape styling Aspose.Cells FillType | group shape placement Aspose.Cells PlacementType | export grouped shapes to XLSX Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new workbook, adds three rectangle shapes named Icon1, Icon2, and Icon3 with distinct fill colors, groups them into a GroupShape named IconSet, configures placement and line formatting for the group, and saves the result as GroupedIcons.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Get the ShapeCollection for the worksheet
        ShapeCollection shapes = worksheet.Shapes;

        // Add three rectangle shapes that will act as icons
        // Parameters: upperLeftRow, top, upperLeftColumn, left, height, width
        Shape icon1 = shapes.AddRectangle(2, 0, 2, 0, 60, 60);
        Shape icon2 = shapes.AddRectangle(2, 0, 5, 0, 60, 60);
        Shape icon3 = shapes.AddRectangle(2, 0, 8, 0, 60, 60);

        // Assign names to the individual shapes
        icon1.Name = "Icon1";
        icon2.Name = "Icon2";
        icon3.Name = "Icon3";

        // Optionally set some visual properties so the icons are distinguishable
        icon1.Fill.FillType = FillType.Solid;
        icon1.Fill.SolidFill.Color = System.Drawing.Color.LightCoral;

        icon2.Fill.FillType = FillType.Solid;
        icon2.Fill.SolidFill.Color = System.Drawing.Color.LightSeaGreen;

        icon3.Fill.FillType = FillType.Solid;
        icon3.Fill.SolidFill.Color = System.Drawing.Color.LightSkyBlue;

        // Group the three shapes together
        Shape[] iconsToGroup = new Shape[] { icon1, icon2, icon3 };
        GroupShape iconSetGroup = shapes.Group(iconsToGroup);

        // Name the group shape
        iconSetGroup.Name = "IconSet";

        // Set a few group properties (optional)
        iconSetGroup.Placement = PlacementType.FreeFloating;
        iconSetGroup.HasLine = true;
        iconSetGroup.LineFormat.ForeColor = System.Drawing.Color.DarkGray;
        iconSetGroup.LineFormat.Weight = 1.5;

        // Save the workbook to a file
        string outputPath = "GroupedIcons.xlsx";
        workbook.Save(outputPath, SaveFormat.Xlsx);

        Console.WriteLine($"Shapes Icon1, Icon2, and Icon3 have been grouped into \"IconSet\" and saved to \"{outputPath}\".");
    }
}
