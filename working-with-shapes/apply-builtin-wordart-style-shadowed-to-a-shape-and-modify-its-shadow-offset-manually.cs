// Title: Adding a WordArt shape with Aspose.Cells C# and understanding the lack of Shadowed style and custom shadow offset support
// Common Searches: how to add WordArt text effect to an Excel worksheet using Aspose.Cells C# | does Aspose.Cells support Shadowed WordArt style for shapes in .NET | setting custom shadow offset for a shape created with Aspose.Cells | limitations of WordArt shadow properties in Aspose.Cells API | example of creating WordArt shape and saving workbook with Aspose.Cells
// Tags: Aspose.Cells add WordArt shape C# | WordArt shadow style limitation Aspose.Cells | custom shape shadow offset .NET | Excel worksheet text effect API | Aspose.Cells shape styling constraints

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // Demonstrates creating a new Workbook, inserting a WordArt shape via the AddTextEffect method, and saving the file. The example also highlights that the current Aspose.Cells version does not provide APIs to apply the built‑in Shadowed WordArt style or to adjust the shadow offset of shapes.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a WordArt shape (text effect) to the worksheet
                // Parameters: preset effect, text, font name, font size, bold, italic,
                // left, top, width, height, rotationAngle, textDirection
                Shape wordArt = worksheet.Shapes.AddTextEffect(
                    MsoPresetTextEffect.TextEffect1,
                    "Aspose.Cells WordArt",
                    "Arial",
                    36,
                    false,
                    false,
                    0,
                    0,
                    400,
                    100,
                    0,   // rotation angle
                    0);  // text direction

                // Note: Advanced WordArt style and shadow settings are not available
                // in the current Aspose.Cells version, so they are omitted.

                // Save the workbook
                string outputPath = "WordArt_Shadowed.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
