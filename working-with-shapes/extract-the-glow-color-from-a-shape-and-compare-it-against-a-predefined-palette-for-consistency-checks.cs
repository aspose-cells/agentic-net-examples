using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class GlowColorCheck
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a rounded rectangle shape to the worksheet
            // Parameters: type, upperLeftRow, upperLeftColumn, top, left, height, width
            Shape shape = worksheet.Shapes.AddAutoShape(
                AutoShapeType.RoundedRectangle, 5, 5, 0, 0, 50, 100);

            // Configure the glow effect with a custom color
            shape.Glow.Size = 10;
            shape.Glow.Color.Color = Color.FromArgb(255, 123, 45, 67); // Custom glow color

            // Extract the glow color from the shape
            CellsColor glowCellsColor = shape.Glow.Color;
            Color glowColor = glowCellsColor.Color;

            // Define a predefined palette of colors for consistency checks
            Color[] predefinedPalette = new Color[]
            {
                Color.Red,
                Color.Green,
                Color.Blue,
                Color.FromArgb(255, 123, 45, 67) // Include the custom color for demonstration
            };

            // Check if the extracted glow color exists in the predefined palette
            bool isInPredefinedPalette = false;
            foreach (Color paletteColor in predefinedPalette)
            {
                if (paletteColor.ToArgb() == glowColor.ToArgb())
                {
                    isInPredefinedPalette = true;
                    break;
                }
            }

            Console.WriteLine($"Extracted glow color: {glowColor}");
            Console.WriteLine($"Is glow color in predefined palette? {isInPredefinedPalette}");

            // Verify if the glow color is present in the workbook's internal palette
            bool isInWorkbookPalette = workbook.IsColorInPalette(glowColor);
            Console.WriteLine($"Is glow color in workbook palette? {isInWorkbookPalette}");

            // If not present, retrieve the closest matching color from the workbook palette
            if (!isInWorkbookPalette)
            {
                Color matchedColor = workbook.GetMatchingColor(glowColor);
                Console.WriteLine($"Closest matching workbook palette color: {matchedColor}");
            }

            // Save the workbook
            string outputPath = "GlowColorCheck.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}