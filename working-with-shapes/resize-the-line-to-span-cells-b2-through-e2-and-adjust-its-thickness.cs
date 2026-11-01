// Title: Resize a line shape to cover cells B2:E2 and set its weight using Aspose.Cells for .NET
// Common Searches: how to make a line shape span B2 to E2 in Aspose.Cells C# | setting line weight for a shape in an Excel file with Aspose.Cells .NET | add a line between two cells and change its thickness using Aspose.Cells API | Aspose.Cells C# resize line shape across a range of cells | adjust line shape thickness programmatically in a workbook with Aspose
// Tags: Aspose.Cells add line shape between specific cells | C# adjust line shape thickness with Aspose.Cells | resize line shape to cover a cell range in Excel using Aspose | line shape weight property Aspose.Cells API | line shape spanning B2:E2 in Aspose.Cells workbook

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a new workbook, adds a line shape from cell B2 to E2, sets its weight to 2 points, and saves the file as LineResized.xlsx.
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

            // Add a line shape from B2 (row 1, column 1) to E2 (row 1, column 4)
            // The overload requires start/end cell positions plus height and width offsets (set to 0)
            Shape shape = sheet.Shapes.AddLine(1, 1, 1, 4, 0, 0);

            // Set line thickness (weight) in points via the Line object's Weight property
            shape.Line.Weight = 2.0f;

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
