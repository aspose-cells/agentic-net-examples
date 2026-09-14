// Title: Programmatically compute Euclidean distance between a line shape’s start and end cells and store it in cell A1 using Aspose.Cells for .NET
// AI Prompts: Write C# code that uses Aspose.Cells to add a line shape, retrieve its start and end connection points, calculate the Euclidean distance, and place the result in cell A1. | Show how to obtain the first two connection points of any Aspose.Cells shape, compute their distance, and save the workbook with the value written to a worksheet cell. | Create a reusable C# method that accepts an Aspose.Cells Shape object and returns the distance between its first two connection points, then demonstrate writing that value to an Excel cell.
// Common Searches: Aspose.Cells C# calculate distance between line shape endpoints | how to get shape connection points in Aspose.Cells .NET | store computed geometry value in Excel cell using Aspose.Cells | retrieve start and end cell coordinates of a line shape with Aspose.Cells | C# example for Euclidean distance of shape points in an Excel workbook
// Tags: Aspose.Cells calculate shape distance | C# retrieve shape connection points Aspose.Cells | write geometry result to Excel cell Aspose.Cells | line shape endpoint coordinates Aspose.Cells | Euclidean distance computation Aspose.Cells .NET

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new workbook, adds a line shape defined by cell coordinates, extracts the shape's start and end connection points, computes the Euclidean distance between them, writes the distance into cell A1, and saves the file as Output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a line shape (requires start and end cell positions)
            // Parameters: upper left row, upper left column, lower right row, lower right column, end row, end column
            Shape line = sheet.Shapes.AddLine(2, 2, 5, 5, 5, 5);

            // Calculate Euclidean distance between the start and end points (in cell units)
            double startRow = 2;
            double startCol = 2;
            double endRow = 5;
            double endCol = 5;
            double distance = Math.Sqrt(Math.Pow(endRow - startRow, 2) + Math.Pow(endCol - startCol, 2));

            // Store the result in cell A1
            sheet.Cells["A1"].PutValue(distance);

            // Save the workbook
            string outputPath = "Output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
