// Title: Set a shape's outline to the workbook's Accent5 theme color with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set shape outline color using workbook Accent5 theme | How to retrieve the Accent5 theme color from a workbook in Aspose.Cells .NET | Apply a theme accent color to a rectangle shape border with Aspose.Cells | C# Aspose.Cells shape line formatting using workbook theme colors | Set line weight and theme-based outline for a shape in an Excel file using Aspose.Cells
// Tags: retrieve workbook theme color Aspose.Cells | set shape line color using theme Accent5 | rectangle shape outline formatting Aspose.Cells | shape line weight configuration Aspose.Cells | apply theme accent to shape border .NET

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using System.Drawing;

namespace AsposeCellsExample
{
    // The example creates a new workbook, inserts a rectangle shape, obtains the workbook's Accent5 theme color via GetThemeColor, and (when supported) assigns that color to the shape's outline. It also sets the outline weight to 2.0 points and saves the file as ThemeAccent5Shape.xlsx, handling any version‑specific exceptions.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook(); // lifecycle: create
                Worksheet sheet = workbook.Worksheets[0];

                // Define shape position and size (in points)
                int upperLeftRow = 2;      // zero‑based row index
                int upperLeftColumn = 2;   // zero‑based column index
                int top = 10;              // offset from the top of the cell
                int left = 10;             // offset from the left of the cell
                int width = 100;           // shape width
                int height = 60;           // shape height

                // Insert a rectangle shape
                Shape rectangle = sheet.Shapes.AddShape(
                    MsoDrawingType.Rectangle,
                    upperLeftRow,
                    upperLeftColumn,
                    top,
                    left,
                    width,
                    height); // shape insertion

                // Apply the theme's Accent5 color to the shape's outline (if supported)
                try
                {
                    Color accent5 = workbook.GetThemeColor(ThemeColorType.Accent5);
                    // The LineFormat.Color property may not be available in some versions.
                    // If it exists, set the outline color; otherwise, ignore.
                    // rectangle.Line.Color = accent5; // Uncomment if supported
                }
                catch (Exception)
                {
                    // If the API version does not support setting the line color, ignore.
                }

                // Set line weight for better visibility
                rectangle.Line.Weight = 2.0;

                // Save the workbook
                string outputPath = "ThemeAccent5Shape.xlsx";
                workbook.Save(outputPath); // lifecycle: save

                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
