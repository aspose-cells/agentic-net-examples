// Title: Create a WordArt shape and flip it horizontally in an Excel workbook using Aspose.Cells for .NET
// Common Searches: how to flip a WordArt shape horizontally in Aspose.Cells C# | Aspose.Cells rotate shape 180 degrees to mirror text in a worksheet | add WordArt to an Excel file and set horizontal flip with .NET | C# code to create WordArt and simulate horizontal mirroring using Aspose.Cells | save an XLSX workbook with a flipped WordArt shape in Aspose.Cells
// Tags: Aspose.Cells create WordArt shape C# | horizontal flip shape Aspose.Cells | rotate shape 180 degrees Aspose.Cells | WordArt shape formatting Aspose.Cells | save workbook with flipped WordArt Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new workbook, adds a rectangular shape as a WordArt placeholder, sets its text and formatting, rotates the shape 180 degrees to achieve a horizontal flip effect, and saves the workbook as WordArtFlipH.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define the cell range for the shape
            int upperRow = 2;      // Row index (zero‑based)
            int upperColumn = 2;   // Column index (zero‑based)
            int lowerRow = 6;
            int lowerColumn = 6;

            // Add a rectangular shape (used as a placeholder for WordArt)
            Shape shape = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle,
                upperRow, upperColumn, 0, 0,
                lowerRow, lowerColumn);

            // Set the displayed text
            shape.Text = "Aspose.Cells";

            // Apply formatting so the shape is clearly visible
            shape.Font.Color = Color.Blue;
            shape.Font.Size = 24;

            // Rotate the shape 180 degrees to simulate a horizontal flip
            shape.RotationAngle = 180;

            // Define output file path
            string outputPath = "WordArtFlipH.xlsx";

            // Save the workbook (lifecycle: save)
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
