// Title: C# – Apply Plastic Material and Soft Lighting (70% intensity) to a Shape with Aspose.Cells
// Description: Creates a workbook, adds a rectangle shape, configures its ThreeDFormat to use Plastic material, applies the Soft lighting preset, sets LightAngle to 70° to simulate 70 % intensity, adds extrusion and a soft‑round bevel, and saves the file as an .xlsx workbook.
// Keywords: Aspose.Cells C# shape 3D | ThreeDFormat plastic material | Soft lighting preset Aspose.Cells | LightAngle intensity | shape extrusion bevel example | Aspose.Cells sample code
// Common Searches: Aspose.Cells set shape material to plastic | How to apply soft lighting to a shape in Aspose.Cells | Set lighting intensity for ThreeDFormat Aspose.Cells | Add extrusion and bevel to a shape using Aspose.Cells | C# example for 3D shape formatting Aspose.Cells
// Developer Intent: Configure a shape’s 3‑D format with Plastic material, Soft lighting, and a 70 % intensity level using Aspose.Cells for .NET.
// Use Cases: Design visually rich dashboard elements with consistent 3‑D styling. | Generate 3‑D icons or call‑outs for presentation‑ready Excel reports. | Automate uniform material and lighting settings across multiple worksheets.
// AI Prompts: Generate C# code that sets a shape's material to Plastic, applies the Soft lighting preset, and uses LightAngle 70° for 70 % intensity with Aspose.Cells. | Show how to add extrusion height and a soft‑round top bevel after configuring lighting on a shape in Aspose.Cells. | Explain how LightAngle values affect perceived lighting intensity in Aspose.Cells ThreeDFormat.

using Aspose.Cells;
using Aspose.Cells.Drawing;

// Create a new workbook and get the first worksheet
Workbook workbook = new Workbook();
Worksheet worksheet = workbook.Worksheets[0];

// Add a rectangle shape to demonstrate 3D formatting
Shape shape = worksheet.Shapes.AddShape(
    MsoDrawingType.Rectangle, // shape type
    1,   // upper left row
    1,   // upper left column
    0,   // offset in pixels from the top of the row
    0,   // offset in pixels from the left of the column
    200, // width in points
    100  // height in points
);

// Access the ThreeDFormat of the shape
ThreeDFormat threeDFormat = shape.ThreeDFormat;

// Set the material to Plastic
threeDFormat.Material = PresetMaterialType.Plastic;

// Apply the Soft lighting preset
threeDFormat.Lighting = LightRigType.Soft;

// Approximate intensity by setting the light angle (0‑359.9 degrees)
// Here 70 degrees is used to represent 70 % intensity as requested
threeDFormat.LightAngle = 70;

// Optional: add some extrusion and bevel to make the 3D effect visible
threeDFormat.ExtrusionHeight = 20;
threeDFormat.TopBevelType = BevelType.SoftRound;
threeDFormat.TopBevelWidth = 10;
threeDFormat.TopBevelHeight = 10;

// Save the workbook
workbook.Save("ThreeDMaterialSoftIntensity.xlsx");
