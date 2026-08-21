// Title: Set Top Lighting Direction for a Rectangle Shape in Aspose.Cells (.NET) – Specular Power Not Supported
// Description: Creates a new workbook, adds a rectangle shape, accesses its ThreeDFormat, sets the lighting direction to Top, notes that the SpecularPower property is unavailable in the current Aspose.Cells API, and saves the workbook as an XLSX file.
// Keywords: Aspose.Cells C# 3D shape lighting | ThreeDFormat LightingDirection Top | Aspose.Cells rectangle shape | SpecularPower property missing | Aspose.Cells .NET example
// Common Searches: how to set lighting direction to top in Aspose.Cells | Aspose.Cells specular power for 3D shapes | C# Aspose.Cells set shape lighting direction | Aspose.Cells ThreeDFormat example | missing SpecularPower in Aspose.Cells
// Developer Intent: Apply a top lighting direction to a rectangle shape and attempt to set a specular power of 80, while handling the fact that the SpecularPower property is not exposed in the Aspose.Cells .NET library.
// Use Cases: Generate a workbook with a 3‑D rectangle shape and configure its lighting direction. | Retrieve and modify ThreeDFormat properties such as lighting direction, depth, or rotation for shapes. | Detect and gracefully handle unsupported 3‑D properties like SpecularPower in Aspose.Cells.
// AI Prompts: Write C# code using Aspose.Cells that sets a shape's LightingDirection to Top and explains the workaround for the missing SpecularPower property. | Suggest techniques to mimic a specular highlight on a 3‑D shape in Aspose.Cells when SpecularPower is unavailable. | Create a try‑catch block that logs a warning if an attempt is made to set an unsupported SpecularPower value on a shape.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a new workbook, adds a rectangle shape, accesses its ThreeDFormat, sets the lighting direction to Top, notes that the SpecularPower property is unavailable in the current Aspose.Cells API, and saves the workbook as an XLSX file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet
            // Parameters: drawing type, upper left row, upper left column, upper left offset X, upper left offset Y, width, height
            Shape shape = worksheet.Shapes.AddShape(MsoDrawingType.Rectangle, 1, 1, 0, 0, 200, 100);
            shape.Text = "3D Shape";

            // Access the shape's 3‑D format
            ThreeDFormat threeDFormat = shape.ThreeDFormat;

            // Set the lighting direction to "Top"
            threeDFormat.LightingDirection = LightRigDirectionType.Top;

            // Note: SpecularPower property is not available in the current Aspose.Cells API.
            // If needed, other 3‑D properties such as Depth, RotationX/Y/Z can be set here.

            // Save the workbook with the applied 3‑D settings
            string outputPath = "Shape3DLightingSpecular.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
