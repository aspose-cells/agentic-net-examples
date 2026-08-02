// Title: Add a Rectangle Shape Linked to Cell A1 in an Excel Workbook with Aspose.Cells for .NET
// Description: Demonstrates how to create a new Workbook, insert a rectangle shape on the first worksheet, bind its text to cell A1 using the LinkedCell property, set an initial value in A1, and save the file as RectangleLinkedCell.xlsx.
// Keywords: Aspose.Cells C# rectangle shape | LinkedCell property | bind shape to cell | dynamic shape text Excel | add shape Aspose.Cells .NET | Excel shape linked cell example
// Common Searches: Aspose.Cells link shape to cell C# | how to bind rectangle text to a worksheet cell | set LinkedCell for a shape in .NET | dynamic text shape Aspose.Cells example
// Developer Intent: Create a workbook, place a rectangle shape, and attach its displayed text to cell A1 so the text updates automatically with the cell value.
// Use Cases: Dashboard labels that always reflect the latest cell values. | Printable forms where shape captions follow user‑entered data. | Report templates with multiple shapes linked to different cells for live data visualization.
// AI Prompts: Generate C# code using Aspose.Cells to add a rectangle shape linked to cell B2, set its size to 150 × 80 points, and write a custom message into B2. | Explain how the LinkedCell property synchronizes a shape's text with a worksheet cell and describe how to refresh the workbook after modifying the cell value.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a new Workbook, insert a rectangle shape on the first worksheet, bind its text to cell A1 using the LinkedCell property, set an initial value in A1, and save the file as RectangleLinkedCell.xlsx.
    public class RectangleShapeLinkedCellDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Get the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a rectangle shape to the worksheet
                // Parameters: upper left row, upper left column, row offset, column offset, width, height
                RectangleShape rectangle = worksheet.Shapes.AddRectangle(1, 1, 0, 0, 120, 60);

                // Link the rectangle's text to cell A1 (dynamic text)
                rectangle.LinkedCell = "$A$1";

                // Set initial text in A1
                worksheet.Cells["A1"].PutValue("Hello, World!");

                // Save the workbook
                string outputPath = "RectangleLinkedCell.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            RectangleShapeLinkedCellDemo.Run();
        }
    }
}
