// Title: How to resize a line shape to cover cells B2‑E2 and set its thickness in Aspose.Cells for .NET
// AI Prompts: Create a new workbook, add a horizontal line shape that starts at cell B2 and ends at cell E2, and set its line weight to 2 points using Aspose.Cells. | Modify an existing line shape so it spans the range B2:E2 and change its thickness to 2 points with the Aspose.Cells .NET API.
// Common Searches: Aspose.Cells .NET resize line shape to specific cell range B2 to E2 | Set line weight for a shape in an Excel file using Aspose.Cells C# | Add horizontal line across multiple cells with Aspose.Cells drawing API | How to adjust line shape dimensions and thickness in a generated workbook with Aspose.Cells
// Tags: add line shape Aspose.Cells .NET | adjust line shape size Aspose.Cells | set line thickness Aspose.Cells C# | horizontal line drawing Aspose.Cells | line shape dimensions Excel Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The program creates a new workbook, inserts a horizontal line shape from cell B2 to E2, sets the line's weight to 2 points, and saves the file as LineResized.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a line shape that starts at cell B2 (row 1, column 1) and ends at cell E2 (row 1, column 4)
            // Height and width are set to 0 for a simple horizontal line
            LineShape line = sheet.Shapes.AddLine(1, 1, 1, 4, 0, 0);

            // Set the line thickness (weight) to 2 points
            line.Line.Weight = 2.0;

            // Save the workbook
            string outputPath = "LineResized.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
