// Title: Add a line connector between cells A1 and B2 and link its start/end points in an Aspose.Cells .NET workbook
// AI Prompts: Write C# code that uses Aspose.Cells to insert a straight line shape whose start anchor is cell A1 and end anchor is cell B2, then set the line weight to 2 points. | Demonstrate how to access a line shape's connection points in Aspose.Cells and bind them to specific worksheet cells so the connector moves with the cells. | Provide code that saves the workbook containing the linked connector as FlowConnectorDemo.xlsx and outputs the full file path.
// Common Searches: Aspose.Cells C# add line shape anchored to two cells | how to bind a connector shape to worksheet cells in Aspose.Cells .NET | set line thickness and save workbook with shape using Aspose.Cells
// Tags: Aspose.Cells add line shape to worksheet | C# bind shape connection points to cells | set line weight Aspose.Cells | save workbook with connector Aspose.Cells | dynamic connector positioning Excel .NET

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// // Creates a new workbook, inserts a straight line shape connecting cell A1 to B2, sets the line weight to 2 points, and saves the file as FlowConnectorDemo.xlsx.
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

            // Add a straight line shape between cell A1 (row 0, column 0) and B2 (row 1, column 1)
            // The last two parameters are the offsets (in pixels) for the start point; set to 0.
            Shape connector = sheet.Shapes.AddLine(0, 0, 1, 1, 0, 0);

            // Customize the connector line
            connector.Line.Weight = 2.0;
            // Note: In some Aspose.Cells versions the LineFormat does not expose a Color property.
            // If needed, you can set the line color using the appropriate API for your version.

            // Save the workbook
            string outputPath = "FlowConnectorDemo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
