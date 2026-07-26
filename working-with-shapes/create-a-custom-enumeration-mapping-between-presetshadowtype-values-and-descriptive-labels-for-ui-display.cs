// Title: C# Mapping Aspose.Cells PresetShadowType Enum to UI‑Friendly Labels
// Description: Learn how to build a static dictionary that links every Aspose.Cells PresetShadowType value to a readable UI string, expose it as a read‑only map, and retrieve labels with a helper method—shown in a shape‑shadow example.
// Keywords: Aspose.Cells PresetShadowType | C# enum to UI label | shadow effect mapping | dictionary lookup Aspose.Cells | shape shadow preset | read‑only dictionary C# | UI friendly enum names | Aspose.Cells example | GitHub Aspose.Cells utilities
// Common Searches: Aspose.Cells map PresetShadowType to text | C# get friendly name for shadow preset | dictionary for PresetShadowType enum | display shadow options in dropdown Aspose.Cells | how to convert PresetShadowType to string
// Developer Intent: Create a reusable lookup that converts PresetShadowType enum values into human‑readable strings for UI components.
// Use Cases: Populate a combo box or dropdown with descriptive shadow names for user selection. | Show the current shadow setting of a shape in a property grid or tooltip. | Validate or reverse‑lookup a user‑chosen label to the corresponding PresetShadowType before applying it.
// AI Prompts: Write a method that returns the PresetShadowType enum value from a given UI label using the existing dictionary. | Generate unit tests for PresetShadowTypeMapper.GetLabel covering all enum values and the unknown case. | Show how to bind PresetShadowTypeMapper.Map to a WinForms ComboBox for selecting shadow presets.

using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Learn how to build a static dictionary that links every Aspose.Cells PresetShadowType value to a readable UI string, expose it as a read‑only map, and retrieve labels with a helper method—shown in a shape‑shadow example.
public static class PresetShadowTypeMapper
{
    // Dictionary that holds the mapping. The keys are the enum values, the values are the UI labels.
    private static readonly Dictionary<PresetShadowType, string> _map = new Dictionary<PresetShadowType, string>
    {
        { PresetShadowType.NoShadow, "No Shadow" },
        { PresetShadowType.Custom, "Custom Shadow" },
        { PresetShadowType.OffsetDiagonalBottomRight, "Outer Shadow – Offset Diagonal Bottom Right" },
        { PresetShadowType.OffsetBottom, "Outer Shadow – Offset Bottom" },
        { PresetShadowType.OffsetDiagonalBottomLeft, "Outer Shadow – Offset Diagonal Bottom Left" },
        { PresetShadowType.OffsetRight, "Outer Shadow – Offset Right" },
        { PresetShadowType.OffsetCenter, "Outer Shadow – Offset Center" },
        { PresetShadowType.OffsetLeft, "Outer Shadow – Offset Left" },
        { PresetShadowType.OffsetDiagonalTopRight, "Outer Shadow – Offset Diagonal Top Right" },
        { PresetShadowType.OffsetTop, "Outer Shadow – Offset Top" },
        { PresetShadowType.OffsetDiagonalTopLeft, "Outer Shadow – Offset Diagonal Top Left" },
        { PresetShadowType.InsideDiagonalTopLeft, "Inner Shadow – Inside Diagonal Top Left" },
        { PresetShadowType.InsideTop, "Inner Shadow – Inside Top" },
        { PresetShadowType.InsideDiagonalTopRight, "Inner Shadow – Inside Diagonal Top Right" },
        { PresetShadowType.InsideLeft, "Inner Shadow – Inside Left" },
        { PresetShadowType.InsideCenter, "Inner Shadow – Inside Center" },
        { PresetShadowType.InsideRight, "Inner Shadow – Inside Right" },
        { PresetShadowType.InsideDiagonalBottomLeft, "Inner Shadow – Inside Diagonal Bottom Left" },
        { PresetShadowType.InsideBottom, "Inner Shadow – Inside Bottom" },
        { PresetShadowType.InsideDiagonalBottomRight, "Inner Shadow – Inside Diagonal Bottom Right" },
        { PresetShadowType.PerspectiveDiagonalUpperLeft, "Outer Shadow – Perspective Diagonal Upper Left" },
        { PresetShadowType.PerspectiveDiagonalUpperRight, "Outer Shadow – Perspective Diagonal Upper Right" },
        { PresetShadowType.Below, "Outer Shadow – Below" },
        { PresetShadowType.PerspectiveDiagonalLowerLeft, "Outer Shadow – Perspective Diagonal Lower Left" },
        { PresetShadowType.PerspectiveDiagonalLowerRight, "Outer Shadow – Perspective Diagonal Lower Right" }
    };

    /// <summary>
    /// Retrieves the UI label for a given <see cref="PresetShadowType"/>.
    /// </summary>
    /// <param name="type">The preset shadow type.</param>
    /// <returns>The descriptive label, or "Unknown" if the type is not in the map.</returns>
    public static string GetLabel(PresetShadowType type)
    {
        return _map.TryGetValue(type, out var label) ? label : "Unknown";
    }

    /// <summary>
    /// Exposes the complete mapping as a read‑only dictionary.
    /// </summary>
    public static IReadOnlyDictionary<PresetShadowType, string> Map => _map;
}

// Demonstration of using the mapper in a typical Aspose.Cells workflow.
class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle: create)
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add a rectangle shape and assign a preset shadow type
        Shape shape = sheet.Shapes.AddRectangle(1, 0, 1, 0, 100, 100);
        shape.ShadowEffect.PresetType = PresetShadowType.OffsetBottom;

        // Obtain the UI‑friendly description for the current shadow type
        string uiLabel = PresetShadowTypeMapper.GetLabel(shape.ShadowEffect.PresetType);
        Console.WriteLine($"Current shadow type: {uiLabel}");

        // Save the workbook (lifecycle: save)
        workbook.Save("ShadowMappingDemo.xlsx");
    }
}
