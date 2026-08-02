// Title: Aspose.Cells .NET – Apply Top Lighting Direction and Specular Power (80) to a Rectangle Shape
// Description: C# example that creates a workbook, inserts a rectangle shape, accesses its ThreeDFormat, sets LightRigDirectionType.Top, notes where a specular property could be set to 80 in newer releases, and saves the file as an Excel workbook.
// Keywords: Aspose.Cells | C# shape lighting | ThreeDFormat | LightRigDirectionType.Top | specular highlight | rectangle shape 3D | Excel automation .NET | 3‑D effects Aspose.Cells | shape formatting | Excel workbook graphics
// Common Searches: Aspose.Cells set shape lighting direction | Top light rig for rectangle shape C# | Specular power 80 Aspose.Cells | ThreeDFormat lighting Aspose.Cells .NET | How to add 3D lighting to Excel shape
// Developer Intent: Configure a rectangle shape in an Excel file to use top‑direction lighting and a specular highlight value of 80 using Aspose.Cells for .NET.
// Use Cases: Enhance dashboard visuals by applying consistent top lighting to chart background shapes. | Prepare presentation‑style Excel reports where shapes need uniform specular highlights for a polished look. | Automate generation of multiple shaped objects with identical 3‑D lighting settings across worksheets.
// AI Prompts: Generate C# code that sets the specular property to 80 on a shape's ThreeDFormat in the latest Aspose.Cells version. | Provide a loop that iterates over all worksheet shapes and applies LightRigDirectionType.Top with specular power 80. | Explain the visual differences between LightRigDirectionType.Top, Front, and Right in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// C# example that creates a workbook, inserts a rectangle shape, accesses its ThreeDFormat, sets LightRigDirectionType.Top, notes where a specular property could be set to 80 in newer releases, and saves the file as an Excel workbook.
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
            // Parameters: drawing type, upper left row, upper left column, top, left, width, height
            Shape shape = worksheet.Shapes.AddShape(MsoDrawingType.Rectangle, 1, 1, 100, 100, 200, 100);

            // Access the 3‑D format of the shape
            ThreeDFormat threeDFormat = shape.ThreeDFormat;

            // Apply three‑dimensional lighting direction "Top"
            threeDFormat.LightingDirection = LightRigDirectionType.Top;

            // Specular highlight setting is not available in this version of Aspose.Cells.
            // If a specular property exists in newer versions, it can be set here, e.g.:
            // threeDFormat.Specular = 80;

            // Save the workbook with the applied 3D settings
            workbook.Save("Shape3DLightingSpecular.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
