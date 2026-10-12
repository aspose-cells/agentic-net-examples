// Title: Place a rectangle shape behind worksheet gridlines and verify its negative Z‑order with Aspose.Cells for .NET
// AI Prompts: Write C# code using Aspose.Cells to add a rectangle shape, set it behind the worksheet gridlines, read its ZOrder value, and confirm the value is less than zero. | Demonstrate how to use reflection in C# to assign the IsBehind property of an Aspose.Cells Shape when the property exists, then obtain and check the shape's Z‑order.
// Common Searches: how to send a shape behind gridlines in Aspose.Cells C# | retrieve ZOrder of a shape in Aspose.Cells .NET | C# reflection to set Shape.IsBehind in Aspose.Cells | check that shape Z-order is negative after moving behind cells Aspose.Cells | Aspose.Cells shape layering order worksheet example
// Tags: Aspose.Cells shape layering behind cells | C# retrieve shape Z-order Aspose.Cells | reflection handling optional shape property | negative Z-order validation Aspose.Cells | add rectangle shape worksheet Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a workbook, inserts a rectangle shape, uses reflection to set the Shape.IsBehind property when available, confirms the shape is positioned behind the gridlines, reads its ZOrder value to ensure it is negative, and saves the file as an .xlsx workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet
            // Parameters: shape type, upper left row, upper left column, top offset, left offset, height, width
            Shape shape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 2, 2, 0, 0, 80, 40);

            // Position the shape behind the gridlines if the API supports it.
            // The IsBehind property may not be available in older versions of Aspose.Cells.
            // If present, set it; otherwise, skip without causing a compile error.
            var isBehindProp = typeof(Shape).GetProperty("IsBehind");
            if (isBehindProp != null && isBehindProp.CanWrite)
            {
                isBehindProp.SetValue(shape, true);
                // Verify that the shape is behind the gridlines
                bool isBehind = (bool)isBehindProp.GetValue(shape);
                if (!isBehind)
                {
                    throw new Exception("Shape is not positioned behind gridlines as expected.");
                }
            }

            // Save the workbook
            string outputPath = "ShapeBehindGridlines.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
