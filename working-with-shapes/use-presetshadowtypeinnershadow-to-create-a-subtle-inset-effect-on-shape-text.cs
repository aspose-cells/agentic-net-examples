// Title: Apply an InsideCenter inner shadow to rectangle shape text in an Excel workbook using Aspose.Cells for .NET
// Common Searches: how to set an inner shadow on shape text with Aspose.Cells C# | Aspose.Cells PresetShadowType.InsideCenter example for a rectangle shape | add inset shadow effect to an Excel shape using Aspose.Cells .NET | C# code to configure shape shadow transparency and blur in Aspose.Cells | create a shape with custom fill, border and inner shadow in an Aspose.Cells workbook
// Tags: shape inset shadow preset | PresetShadowType.InsideCenter example | shape shadow properties Aspose.Cells | Excel rectangle shape text styling | C# Aspose.Cells shape fill line configuration

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsInnerShadowDemo
{
    // The sample creates a new workbook, inserts a rectangle shape with a light‑yellow fill and dark‑gray border, assigns text to the shape, and applies an inner (InsideCenter) shadow using PresetShadowType.InsideCenter. The shadow is configured with 50% transparency, a slight blur, a small distance, and a black color before the workbook is saved as an XLSX file.
    class Program
    {
        static void Main(string[] args)
        {
            // Path where the output workbook will be saved
            const string outputPath = "InnerShadowShapeTextDemo.xlsx";

            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Add a rectangle shape that will contain text
                // Parameters: upper left row, upper left column, upper left offset X, upper left offset Y, width, height
                Shape shape = sheet.Shapes.AddRectangle(2, 1, 0, 0, 300, 100);

                // Set the shape's fill and line to make the inset effect more visible
                shape.FillFormat.ForeColor = Color.LightYellow;
                shape.LineFormat.Weight = 0.5;
                shape.LineFormat.ForeColor = Color.DarkGray;

                // Add text to the shape
                shape.Text = "Inner Shadow Inset Effect";

                // Configure the shadow effect to use an inner shadow (InsideCenter)
                ShadowEffect shadow = shape.ShadowEffect;
                shadow.PresetType = PresetShadowType.InsideCenter; // subtle inset
                shadow.Transparency = 0.5;   // 50% transparent for a gentle look
                shadow.Blur = 5;            // slight blur
                shadow.Distance = 2;        // small distance to keep it inset
                shadow.Angle = 0;           // angle is ignored for inner shadows but set to default
                shadow.Size = 0;            // size is meaningless for inner shadows
                // Set a dark color for the inner shadow
                CellsColor shadowColor = workbook.CreateCellsColor();
                shadowColor.Color = Color.Black;
                shadow.Color = shadowColor;

                // Save the workbook
                workbook.Save(outputPath, SaveFormat.Xlsx);

                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
                Console.WriteLine("Shape details:");
                Console.WriteLine($"  Text: {shape.Text}");
                Console.WriteLine($"  Shadow PresetType: {shadow.PresetType}");
                Console.WriteLine($"  Transparency: {shadow.Transparency}");
                Console.WriteLine($"  Blur: {shadow.Blur}");
                Console.WriteLine($"  Distance: {shadow.Distance}");
            }
            catch (Exception ex)
            {
                // Output any errors that occur during processing
                Console.WriteLine("An error occurred:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
