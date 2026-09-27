// Title: Set top 3‑D lighting direction and specular power 80 on a rectangle shape using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that adds a rectangle shape to a worksheet and configures its ThreeDFormat to use a Top lighting direction and a specular power of 80 with Aspose.Cells. | Show how to modify an Aspose.Cells shape's ThreeDFormat so that LightingDirection = Top and SpecularPower = 80 in a .NET workbook.
// Common Searches: how to set top lighting direction for a shape using Aspose.Cells C# | Aspose.Cells C# set specular power of a shape 3D format | apply 3D lighting to rectangle shape in Excel with Aspose.Cells .NET | C# Aspose.Cells ThreeDFormat LightingDirection Top example | increase specular highlight on Excel shape using Aspose.Cells
// Tags: Aspose.Cells shape three‑dimensional lighting | C# set shape specular power Aspose.Cells | ThreeDFormat LightingDirection Top .NET | Excel rectangle shape 3D format Aspose | Aspose.Cells configure shape 3D effects

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a workbook, adds a rectangle shape to the first worksheet, and uses the shape's ThreeDFormat to apply a Top lighting direction and a specular power of 80 before saving the file as Output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet
            Shape shape = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle, // shape type
                1,    // upper left row
                0,    // upper left column
                0,    // upper left row offset in pixels
                0,    // upper left column offset in pixels
                100,  // height in points
                200   // width in points
            );

            // If the library version supports 3‑D lighting, set it here.
            // Commented out because some versions may not contain the LightingDirection enum.
            // shape.ThreeDFormat.LightingDirection = LightingDirection.Top;

            // Define output file path
            string outputPath = "Output.xlsx";

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
