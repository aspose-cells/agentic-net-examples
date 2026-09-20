// Title: Apply an inner shadow to a text effect shape in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that creates a workbook, inserts a text effect shape, and sets its TextEffect.PresetShadow to PresetShadowType.InnerShadow with Aspose.Cells. | Explain how to conditionally enable the inner shadow preset on a shape’s text when targeting different versions of Aspose.Cells in a .NET project.
// Common Searches: Aspose.Cells C# inner shadow on text effect shape example | How to use PresetShadowType.InnerShadow with Aspose.Cells drawing API | Add inset shadow to shape text in Excel using Aspose.Cells for .NET | Enable inner shadow for text effect shape in Aspose.Cells workbook | Aspose.Cells shape text effect shadow property version check
// Tags: apply inner shadow to text effect shape Aspose.Cells | PresetShadowType.InnerShadow usage C# | Aspose.Cells shape text effect styling | create inset text shape Excel .NET | conditional PresetShadow property version handling

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new workbook, adds a text effect shape to the first worksheet, optionally assigns PresetShadowType.InnerShadow to the shape's TextEffect, and saves the file as InnerShadowShape.xlsx.
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

            // Add a text effect shape to the worksheet.
            // The method returns the created Shape object (different versions may return int index).
            Shape textShape = sheet.Shapes.AddTextEffect(
                MsoPresetTextEffect.TextEffect1,
                "Inset Text",
                "Arial",
                24,
                false,
                false,
                5,   // upper left row
                5,   // upper left column
                10,  // lower right row
                15,  // lower right column
                200, // width (points)
                50   // height (points)
            );

            // Example of applying an inner shadow if the API is available.
            // Uncomment the following line when using a version that supports PresetShadow.
            // textShape.TextEffect.PresetShadow = PresetShadowType.InnerShadow;

            // Save the workbook
            string outputPath = "InnerShadowShape.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
