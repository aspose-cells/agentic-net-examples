using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class Program
{
    static void Main()
    {
        // Mapping between PresetShadowType enum values and user‑friendly descriptions.
        var shadowTypeLabels = new Dictionary<PresetShadowType, string>
        {
            { PresetShadowType.NoShadow, "No shadow" },
            { PresetShadowType.Custom, "Custom shadow" },
            { PresetShadowType.OffsetDiagonalBottomRight, "Outer shadow offset diagonal bottom right" },
            { PresetShadowType.OffsetBottom, "Outer shadow offset bottom" },
            { PresetShadowType.OffsetDiagonalBottomLeft, "Outer shadow offset diagonal bottom left" },
            { PresetShadowType.OffsetRight, "Outer shadow offset right" },
            { PresetShadowType.OffsetCenter, "Outer shadow offset center" },
            { PresetShadowType.OffsetLeft, "Outer shadow offset left" },
            { PresetShadowType.OffsetDiagonalTopRight, "Outer shadow offset diagonal top right" },
            { PresetShadowType.OffsetTop, "Outer shadow offset top" },
            { PresetShadowType.OffsetDiagonalTopLeft, "Outer shadow offset diagonal top left" },
            { PresetShadowType.InsideDiagonalTopLeft, "Inner shadow inside diagonal top left" },
            { PresetShadowType.InsideTop, "Inner shadow inside top" },
            { PresetShadowType.InsideDiagonalTopRight, "Inner shadow inside diagonal top right" },
            { PresetShadowType.InsideLeft, "Inner shadow inside left" },
            { PresetShadowType.InsideCenter, "Inner shadow inside center" },
            { PresetShadowType.InsideRight, "Inner shadow inside right" },
            { PresetShadowType.InsideDiagonalBottomLeft, "Inner shadow inside diagonal bottom left" },
            { PresetShadowType.InsideBottom, "Inner shadow inside bottom" },
            { PresetShadowType.InsideDiagonalBottomRight, "Inner shadow inside diagonal bottom right" },
            { PresetShadowType.PerspectiveDiagonalUpperLeft, "Outer shadow perspective diagonal upper left" },
            { PresetShadowType.PerspectiveDiagonalUpperRight, "Outer shadow perspective diagonal upper right" },
            { PresetShadowType.Below, "Outer shadow below" },
            { PresetShadowType.PerspectiveDiagonalLowerLeft, "Outer shadow perspective diagonal lower left" },
            { PresetShadowType.PerspectiveDiagonalLowerRight, "Outer shadow perspective diagonal lower right" }
        };

        // Create a workbook and a shape to demonstrate applying a shadow type.
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Shape shape = worksheet.Shapes.AddRectangle(1, 0, 1, 0, 100, 100);

        // Set a preset shadow type.
        shape.ShadowEffect.PresetType = PresetShadowType.OffsetDiagonalBottomRight;

        // Retrieve and display the friendly label for the current preset type.
        string label = shadowTypeLabels[shape.ShadowEffect.PresetType];
        Console.WriteLine($"Current shadow type: {shape.ShadowEffect.PresetType} - {label}");

        // Save the workbook.
        workbook.Save("ShadowMappingDemo.xlsx");
    }
}