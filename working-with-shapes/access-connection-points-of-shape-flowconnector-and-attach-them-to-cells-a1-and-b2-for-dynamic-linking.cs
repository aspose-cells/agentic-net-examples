// Title: Position a FlowConnector line shape between cells A1 and B2 with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set FlowConnector start cell A1 and end cell B2 | How to attach a line connector to specific worksheet cells using Aspose.Cells | C# code to bind a FlowConnector shape to Excel cells dynamically | Using Aspose.Cells API to align connector shape with cells A1 and B2 | Set UpperLeftRow UpperLeftColumn LowerRightRow LowerRightColumn for line shape in Aspose.Cells
// Tags: flowconnector connection points Aspose.Cells | line shape positioning worksheet cells C# | connector shape cell anchoring Aspose.Cells | Aspose.Cells shape coordinate properties | dynamic cell‑linked connector Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new workbook, adds a straight line (FlowConnector) shape, sets its UpperLeftRow/Column to cell A1 and LowerRightRow/Column to cell B2, and saves the file as 'FlowConnector.xlsx', demonstrating how to anchor a connector shape to specific cells using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a straight line shape (acts as a connector) to the worksheet
            // Parameters: upper left row, upper left column, lower right row, lower right column, height, width
            Shape line = sheet.Shapes.AddLine(0, 0, 1, 1, 0, 0);

            // Position the line: start at cell A1 and end at cell B2
            line.UpperLeftRow = 0;
            line.UpperLeftColumn = 0;
            line.LowerRightRow = 1;
            line.LowerRightColumn = 1;

            // Save the workbook
            string outputPath = "FlowConnector.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
