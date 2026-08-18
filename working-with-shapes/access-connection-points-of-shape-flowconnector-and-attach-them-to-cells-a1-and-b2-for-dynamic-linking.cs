// Title: C# – Link FlowConnector Shape Connection Points to Cells A1 and B2 with Aspose.Cells
// Description: Demonstrates how to create a workbook, add line shapes as FlowConnector objects, retrieve their connection points, and attach the first connector's start point to cell A1 and a second connector's start point to cell B2 using SetLinkedCell. The linked connectors move automatically with the referenced cells, and the workbook is saved as an Excel file.
// Keywords: Aspose.Cells C# example | FlowConnector SetLinkedCell | link shape to cell Aspose.Cells | retrieve connection points line shape | dynamic connector positioning Excel | Aspose.Cells API GetConnectionPoints | Excel shape linking .NET | code sample GitHub Aspose.Cells
// Common Searches: how to link a FlowConnector to a cell using Aspose.Cells | Aspose.Cells GetConnectionPoints C# | SetLinkedCell for line shape in Excel workbook | dynamic shape positioning with Aspose.Cells | attach multiple connectors to different cells programmatically
// Developer Intent: Attach FlowConnector shape connection points to specific worksheet cells so the connectors adjust automatically when the cells move.
// Use Cases: Create a workbook, add a line shape, and bind its start point to cell A1 for responsive diagram layouts. | Retrieve and log X/Y coordinates of a connector's connection points for layout verification or custom calculations. | Add a second connector, link its start point to cell B2, and save the workbook with both linked shapes.
// AI Prompts: Generate C# code that links both ends of a FlowConnector to two separate cells using Aspose.Cells. | Explain how to refresh linked cells of existing connectors after inserting or deleting rows or columns. | Show how to map GetConnectionPoints values to worksheet cell addresses and convert between pixel and cell units.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to create a workbook, add line shapes as FlowConnector objects, retrieve their connection points, and attach the first connector's start point to cell A1 and a second connector's start point to cell B2 using SetLinkedCell. The linked connectors move automatically with the referenced cells, and the workbook is saved as an Excel file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a line shape (used as a simple connector)
            // Parameters: upper left row, upper left column, lower right row, lower right column, upper left row offset, upper left column offset
            Shape flowConnector = worksheet.Shapes.AddLine(2, 2, 8, 8, 0, 0);

            // Retrieve the connection points of the shape
            float[][] connectionPoints = flowConnector.GetConnectionPoints();

            // Output the connection points to the console
            Console.WriteLine("Connection Points of FlowConnector:");
            for (int i = 0; i < connectionPoints.Length; i++)
            {
                Console.WriteLine($"Point {i + 1}: X = {connectionPoints[i][0]}, Y = {connectionPoints[i][1]}");
            }

            // Dynamically link the first connection point to cell A1
            flowConnector.SetLinkedCell("$A$1", false, true);

            // Add a second line shape and link its first connection point to cell B2
            Shape secondConnector = worksheet.Shapes.AddLine(10, 2, 16, 8, 0, 0);
            secondConnector.SetLinkedCell("$B$2", false, true);

            // Save the workbook with the linked connectors
            string outputPath = "FlowConnectorLinked.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
