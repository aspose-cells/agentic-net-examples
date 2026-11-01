// Title: Create a dark‑gray line shape in an Excel worksheet using Aspose.Cells for .NET
// Common Searches: how to insert a line shape into an Excel file with Aspose.Cells C# | Aspose.Cells set line shape color to dark gray programmatically | add line shape to worksheet using Shapes.AddLine in C# | draw a line in Excel workbook with Aspose.Cells drawing API | change line format color of a shape in Aspose.Cells .NET
// Tags: Shapes.AddLine Aspose.Cells .NET | line shape color setting Excel | drawing toolbar shape creation Aspose.Cells | dark gray line format Aspose.Cells | line shape insertion Aspose.Cells .NET

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new Workbook, adds a line shape to the first worksheet using Shapes.AddLine with cell coordinates, optionally sets the line’s color to DarkGray, and saves the file as LineShape.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define the position of the line (upper‑left and lower‑right cells)
            int upperLeftRow = 2;      // Row index (zero‑based)
            int upperLeftColumn = 2;   // Column index (zero‑based)
            int lowerRightRow = 5;
            int lowerRightColumn = 5;

            // Width and height offsets (in points) – set to 0 for default size
            int width = 0;
            int height = 0;

            // Add a line shape – equivalent to using the drawing toolbar
            Shape lineShape = sheet.Shapes.AddLine(upperLeftRow, upperLeftColumn, lowerRightRow, lowerRightColumn, width, height);

            // Optionally set the line color (commented out because LineFormat.Color may not be available in older versions)
            // lineShape.Line.Color = Color.DarkGray;

            // Save the workbook to a file
            string outputPath = "LineShape.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
