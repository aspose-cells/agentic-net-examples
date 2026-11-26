// Title: Add a rectangle shape with a solid RGB fill color to an Excel worksheet using Aspose.Cells for .NET
// Common Searches: how to add a rectangle shape with solid RGB fill in Aspose.Cells C# | Aspose.Cells set shape fill type to solid and specify color | C# create rectangle shape in Excel and apply red fill using Aspose.Cells | Aspose.Cells drawing API change shape fill color programmatically | save workbook after adding colored shape with Aspose.Cells .NET
// Tags: add rectangle shape Aspose.Cells .NET | set shape solid fill Aspose.Cells | apply RGB color to shape Aspose.Cells | Aspose.Cells shape drawing API | save workbook with shape Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new Workbook, inserts a rectangle shape on the first worksheet, sets its FillType to Solid, applies a red RGB color via Color.FromArgb, optionally configures the shape border, and saves the file as an XLSX workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a rectangle shape.
            // Parameters: upperLeftRow, upperLeftColumn, top (pixel offset), left (pixel offset), height (pixel), width (pixel)
            Shape rectangle = worksheet.Shapes.AddRectangle(1, 1, 0, 0, 100, 150);

            // Apply a solid fill using an RGB color (example: pure red = 255,0,0)
            rectangle.Fill.FillType = FillType.Solid;
            rectangle.Fill.SolidFill.Color = Color.FromArgb(255, 0, 0); // Red

            // Optional: set a line color so the shape border is visible
            // Note: In some Aspose.Cells versions the Line property may not expose a Color setter.
            // If needed, uncomment the following line after verifying the API version.
            // rectangle.Line.Color = Color.Black;

            // Save the workbook to a file
            string outputPath = "ShapeSolidFill.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
