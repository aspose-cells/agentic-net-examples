using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

namespace AsposeCellsDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Add a rectangle shape (as a substitute for SmartArt) to the worksheet
                // Parameters: upper left row, upper left column, lower right row, lower right column, height, width
                Shape shape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 1, 1, 5, 5, 100, 100);

                // Link the shape to a specific cell (e.g., C5)
                shape.LinkedCell = "$C$5";

                // Save the workbook
                string outputFile = "SmartArtLinkedCell.xlsx";
                workbook.Save(outputFile);
                Console.WriteLine($"Workbook saved successfully: {Path.GetFullPath(outputFile)}");
            }
            catch (Exception ex)
            {
                // Handle any runtime errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}