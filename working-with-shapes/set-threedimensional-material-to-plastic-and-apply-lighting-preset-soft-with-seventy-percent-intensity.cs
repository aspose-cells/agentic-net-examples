// Title: C# Example: Set Shape 3‑D Material to Plastic and Apply Soft Lighting (70% Intensity) with Aspose.Cells
// Description: Demonstrates how to create a workbook, add a rectangle shape, configure its ThreeDFormat to use the Plastic material, apply the Soft lighting preset, approximate 70 % lighting intensity via LightAngle, add extrusion and bevel for depth, and save the file as ThreeDMaterialSoftLighting.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | .NET | shape 3D material | Plastic material | Soft lighting | LightAngle intensity | ThreeDFormat | extrusion | bevel | rectangle shape | code example | tutorial
// Common Searches: Aspose.Cells set shape material to plastic | apply soft lighting to a shape in Aspose.Cells .NET | control lighting intensity for 3D shapes Aspose.Cells | add extrusion and bevel to a shape using Aspose.Cells | C# 3D shape formatting example Aspose.Cells
// Developer Intent: Apply Plastic material and Soft lighting (≈70 % intensity) to a shape via Aspose.Cells for .NET.
// Use Cases: Design a highlighted call‑out box with plastic finish and soft lighting for a financial dashboard. | Create a 3‑D button with extrusion and subtle lighting for an interactive spreadsheet UI. | Style chart legends or legends in reports using plastic material and soft lighting to improve visual hierarchy.
// AI Prompts: Generate C# code with Aspose.Cells that sets a shape’s material to Plastic, applies the Soft lighting preset at 70 % intensity, and adds extrusion and bevel settings. | Explain how LightAngle influences perceived lighting intensity in Aspose.Cells 3‑D formatting and suggest alternative properties for finer control. | Provide a step‑by‑step guide to apply different 3‑D materials and lighting presets to multiple shapes in a workbook using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to create a workbook, add a rectangle shape, configure its ThreeDFormat to use the Plastic material, apply the Soft lighting preset, approximate 70 % lighting intensity via LightAngle, add extrusion and bevel for depth, and save the file as ThreeDMaterialSoftLighting.xlsx using Aspose.Cells for .NET.
class Set3DMaterialAndLighting
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a shape (e.g., a rectangle) to demonstrate 3‑D formatting
        Shape shape = worksheet.Shapes.AddShape(MsoDrawingType.Rectangle, 1, 1, 0, 0, 200, 100);
        shape.Text = "3D Plastic with Soft Lighting";

        // Access the ThreeDFormat of the shape
        ThreeDFormat threeD = shape.ThreeDFormat;

        // Set the material to Plastic
        threeD.Material = PresetMaterialType.Plastic;

        // Apply the Soft lighting preset
        threeD.Lighting = LightRigType.Soft;

        // Approximate intensity by setting the LightAngle (0‑359.9 degrees);
        // using 70 as a representation of 70 % intensity
        threeD.LightAngle = 70;

        // Optional: add some extrusion so the 3‑D effect is visible
        threeD.ExtrusionHeight = 20;
        threeD.TopBevelType = BevelType.SoftRound;
        threeD.TopBevelWidth = 10;
        threeD.TopBevelHeight = 10;

        // Save the workbook
        workbook.Save("ThreeDMaterialSoftLighting.xlsx");
    }
}
