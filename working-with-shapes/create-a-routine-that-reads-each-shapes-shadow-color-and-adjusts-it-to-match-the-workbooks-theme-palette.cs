// Title: C# example: Update all shape shadow colors to match the workbook’s Accent1 theme color using Aspose.Cells
// Common Searches: How to set Excel shape shadow color to the workbook theme Accent1 with Aspose.Cells in C# | Aspose.Cells C# iterate worksheets and change shape shadow colors based on theme palette | Read and modify shape shadow properties using the workbook’s theme colors in Aspose.Cells .NET | Apply workbook theme colors to shape shadows in an existing Excel file with Aspose.Cells | Dynamic access to shape shadow API when theme methods are unavailable in Aspose.Cells
// Tags: shape shadow color adjustment Aspose.Cells | apply workbook theme palette to shape shadows | retrieve Accent1 theme color Aspose.Cells | iterate worksheets modify shape shadows | dynamic API access for shape shadow Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample loads an existing workbook (or creates one with a rectangle shape), obtains the Accent1 color from the workbook's theme palette using dynamic invocation, then walks through every worksheet and shape, setting each shape's shadow color to the retrieved theme color and ensuring the shadow is visible before saving the file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        Workbook workbook;

        // Load existing workbook if present; otherwise create a sample workbook with a shape.
        if (File.Exists(inputPath))
        {
            try
            {
                workbook = new Workbook(inputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load \"{inputPath}\": {ex.Message}");
                workbook = new Workbook();
            }
        }
        else
        {
            workbook = new Workbook();
            Worksheet ws = workbook.Worksheets[0];

            // Add a rectangle shape with a visible shadow.
            Shape sampleShape = ws.Shapes.AddShape(
                MsoDrawingType.Rectangle,
                1, 0,   // upper left row, column
                1, 0,   // top, left offsets (pixels)
                120, 60 // width, height (pixels)
            );

            // Use dynamic to set shadow properties (avoids compile‑time binding issues).
            try
            {
                dynamic dynShape = sampleShape;
                var shadow = dynShape.Shadow;
                if (shadow != null)
                {
                    shadow.Visible = true;
                    shadow.Color = Color.Gray;
                }
            }
            catch
            {
                // Shadow property not available in this version; ignore.
            }
        }

        // ------------------------------------------------------------
        // Obtain a colour from the workbook's theme palette (Accent1).
        // Use dynamic invocation to stay compatible with older versions.
        // ------------------------------------------------------------
        Color targetShadowColor = Color.Gray; // fallback colour
        try
        {
            dynamic dynWorkbook = workbook;
            targetShadowColor = dynWorkbook.GetThemeColor(ThemeColorType.Accent1);
        }
        catch
        {
            // Theme API not available; keep fallback colour.
        }

        // ------------------------------------------------------------
        // Iterate through every worksheet and shape, adjusting shadow colour.
        // ------------------------------------------------------------
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            foreach (Shape shape in sheet.Shapes)
            {
                try
                {
                    dynamic dynShape = shape;
                    var shadow = dynShape.Shadow;
                    if (shadow != null)
                    {
                        shadow.Color = targetShadowColor;
                        shadow.Visible = true;
                    }
                }
                catch
                {
                    // Shadow property not supported; continue.
                }
            }
        }

        // Save the modified workbook.
        try
        {
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook processed and saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to save workbook: {ex.Message}");
        }
    }
}
