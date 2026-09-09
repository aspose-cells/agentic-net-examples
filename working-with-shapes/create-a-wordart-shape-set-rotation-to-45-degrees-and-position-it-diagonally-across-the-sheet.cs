// Title: Add a rotated WordArt shape spanning cells A1 to J20 in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Insert a WordArt object with custom text, set its RotationAngle to 45°, and anchor it from cell A1 to J20 using the Aspose.Cells Shapes API in C#. | Apply a solid light‑blue fill and a dark‑blue line border to the WordArt shape, employing reflection to handle differing Aspose.Cells version property names.
// Common Searches: C# Aspose.Cells rotate WordArt 45 degrees | how to set RotationAngle for WordArt in Aspose.Cells C# | Aspose.Cells place WordArt across a range of cells in Excel | apply fill color to WordArt shape using Aspose.Cells .NET | add line border to WordArt with Aspose.Cells handling version differences | save workbook with rotated WordArt using Aspose.Cells C#
// Tags: Aspose.Cells add WordArt shape C# | rotate WordArt shape Aspose.Cells | anchor WordArt to cell range Aspose.Cells | solid fill WordArt Aspose.Cells | line border WordArt Aspose.Cells | reflection compatibility Aspose.Cells shape formatting

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new workbook, adds a WordArt shape that spans cells A1 to J20, rotates it 45°, applies a light‑blue solid fill and a dark‑blue line border (using reflection for version‑agnostic property access), and saves the file as WordArtDemo.xlsx.
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

            // Define the position of the WordArt (from top‑left cell A1 to bottom‑right cell J20)
            int upperLeftRow = 0;          // Row 0 (A)
            int upperLeftColumn = 0;       // Column 0 (1)
            int lowerRightRow = 20;        // Row 20
            int lowerRightColumn = 9;      // Column J (0‑based index)

            // Width and height of the shape (in points)
            int shapeWidth = 400;
            int shapeHeight = 100;

            // Add a WordArt shape with the specified text
            Shape wordArt = sheet.Shapes.AddWordArt(
                PresetWordArtStyle.WordArtStyle1,
                "Aspose.Cells WordArt",
                upperLeftRow,
                upperLeftColumn,
                lowerRightRow,
                lowerRightColumn,
                shapeWidth,
                shapeHeight);

            // Set rotation to 45 degrees
            wordArt.RotationAngle = 45;

            // Apply fill formatting (if supported by the current Aspose.Cells version)
            try
            {
                wordArt.Fill.FillType = FillType.Solid;
                // Some versions expose SolidFillColor; others use SolidFillColor property directly.
                // Use reflection as a fallback to maintain compatibility.
                var solidFillProp = wordArt.Fill.GetType().GetProperty("SolidFillColor");
                if (solidFillProp != null && solidFillProp.CanWrite)
                {
                    solidFillProp.SetValue(wordArt.Fill, Color.LightBlue);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Fill formatting not applied: {ex.Message}");
            }

            // Apply line formatting (if supported by the current Aspose.Cells version)
            try
            {
                // Some versions expose Line.Fill; others expose Line.FillFormat.
                var lineFillProp = wordArt.Line.GetType().GetProperty("Fill");
                if (lineFillProp != null)
                {
                    var lineFill = lineFillProp.GetValue(wordArt.Line);
                    var fillTypeProp = lineFill.GetType().GetProperty("FillType");
                    var solidFillProp = lineFill.GetType().GetProperty("SolidFillColor");
                    if (fillTypeProp != null && solidFillProp != null)
                    {
                        fillTypeProp.SetValue(lineFill, FillType.Solid);
                        solidFillProp.SetValue(lineFill, Color.DarkBlue);
                    }
                }
                else
                {
                    // Fallback for older APIs
                    var lineFillFormatProp = wordArt.Line.GetType().GetProperty("FillFormat");
                    if (lineFillFormatProp != null)
                    {
                        var lineFill = lineFillFormatProp.GetValue(wordArt.Line);
                        var fillTypeProp = lineFill.GetType().GetProperty("FillType");
                        var solidFillProp = lineFill.GetType().GetProperty("SolidFillColor");
                        if (fillTypeProp != null && solidFillProp != null)
                        {
                            fillTypeProp.SetValue(lineFill, FillType.Solid);
                            solidFillProp.SetValue(lineFill, Color.DarkBlue);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Line formatting not applied: {ex.Message}");
            }

            // Save the workbook
            workbook.Save("WordArtDemo.xlsx", SaveFormat.Xlsx);
            Console.WriteLine("Workbook saved successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
