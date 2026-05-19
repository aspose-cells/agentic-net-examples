using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet (using positional arguments)
            // Parameters: upperLeftRow, upperLeftColumn, upperLeftRowOffset, upperLeftColumnOffset, height, width
            Shape shape = worksheet.Shapes.AddRectangle(1, 0, 0, 0, 100, 200);

            // Retrieve the connection points of the shape
            float[][] points = shape.GetConnectionPoints();

            // Generate a CSV file with point index and coordinates
            string csvFile = "ConnectionPoints.csv";
            using (StreamWriter writer = new StreamWriter(csvFile))
            {
                writer.WriteLine("Index,X,Y"); // Header
                for (int i = 0; i < points.Length; i++)
                {
                    writer.WriteLine($"{i},{points[i][0]},{points[i][1]}");
                }
            }

            // Save the workbook (optional, to keep the shape in the file)
            string workbookFile = "ShapeWithConnectionPoints.xlsx";
            workbook.Save(workbookFile);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}