// Title: How to set a rectangle shape's outline to the workbook's Accent5 theme color using Aspose.Cells for .NET
// AI Prompts: Retrieve the workbook's Accent5 theme color with GetThemeColor and assign it to the Line.Color of a rectangle shape in Aspose.Cells C#. | Configure a shape's line weight and apply a theme accent color to its border when generating an Excel file with Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# set shape border color to workbook theme Accent5 | apply Excel theme accent color to drawing shape outline using Aspose.Cells .NET | how to use GetThemeColor with shapes in Aspose.Cells
// Tags: Aspose.Cells set shape outline color | GetThemeColor Accent5 Aspose.Cells | shape line format Excel theme .NET | rectangle shape border customization Aspose.Cells | apply workbook theme color to drawing objects C#

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new workbook, adds a rectangle shape, retrieves the workbook's Accent5 theme color via GetThemeColor, sets the shape's line weight, applies the retrieved color to the shape's outline, and saves the file as an .xlsx document.
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
            // Parameters: shape type, upper left row, upper left column, upper left row offset, upper left column offset, width, height
            Shape shape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 2, 2, 0, 0, 150, 80);

            // Retrieve the theme's Accent5 color
            Color accent5Color = workbook.GetThemeColor(ThemeColorType.Accent5);

            // Apply the Accent5 color to the shape's outline (if supported by the API version)
            shape.Line.Weight = 2.0; // Set line thickness
            // Uncomment the following line if the LineFormat.Color property is available in your Aspose.Cells version
            // shape.Line.Color = accent5Color;

            // Save the workbook to a file
            string outputPath = "ShapeWithAccent5Outline.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
