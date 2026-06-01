using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsFlowConnectorDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a straight line shape (used as a connector) to the worksheet
                // Parameters: upper left row, upper left column, upper left pixel offset,
                // lower right row, lower right column, lower right pixel offset
                Shape flowConnector = worksheet.Shapes.AddLine(
                    2, 2, 0,   // start at cell C3 (row 2, column 2)
                    5, 5, 0);  // end at cell F6 (row 5, column 5)

                // Retrieve the connection points of the shape
                float[][] connectionPoints = flowConnector.GetConnectionPoints();

                // Output the connection points for verification
                Console.WriteLine("Connector Connection Points:");
                for (int i = 0; i < connectionPoints.Length; i++)
                {
                    Console.WriteLine($"Point {i + 1}: X = {connectionPoints[i][0]}, Y = {connectionPoints[i][1]}");
                }

                // Link the shape to a worksheet cell (example)
                flowConnector.SetLinkedCell("$A$1", false, true);

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
}