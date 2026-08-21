// Title: Aspose.Cells for .NET – Load Workbook, Select Worksheet, and Add a Rectangle Shape
// Description: C# example that loads an existing Excel file (input.xlsx) with Aspose.Cells, selects the first worksheet, inserts a rectangle shape at row 2/column 2, sets its name and light‑blue fill, and saves the result to output.xlsx.
// Keywords: Aspose.Cells C# load workbook | select worksheet Aspose.Cells | add rectangle shape .NET | shape fill color Aspose.Cells | Aspose.Cells drawing API | Excel automation C# | save workbook after shape edit | Aspose.Cells example | programmatic Excel shapes
// Common Searches: how to add a rectangle shape to an Excel sheet using Aspose.Cells | load workbook and select worksheet Aspose.Cells C# | set shape fill color with Aspose.Cells .NET | Aspose.Cells add shape and save workbook | C# code for drawing shapes in Excel with Aspose
// Developer Intent: Load an existing Excel file, choose a worksheet, insert a rectangle shape, customize its appearance, and save the modified workbook using Aspose.Cells for .NET.
// Use Cases: Highlight a data block in a report by drawing a colored rectangle. | Create a placeholder shape for a chart or image in a template workbook. | Add named shapes that can be referenced later in automated processing or reporting pipelines.
// AI Prompts: Generate C# code that loads a workbook, selects the second worksheet, and adds an ellipse shape with a red border using Aspose.Cells. | Explain how to change the fill color of an existing shape in a saved Excel file with Aspose.Cells for .NET. | Show how to iterate over all shapes in a worksheet and modify their properties with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShapeDemo
{
    // C# example that loads an existing Excel file (input.xlsx) with Aspose.Cells, selects the first worksheet, inserts a rectangle shape at row 2/column 2, sets its name and light‑blue fill, and saves the result to output.xlsx.
    class Program
    {
        static void Main()
        {
            // Load an existing workbook from file (uses Workbook(string) constructor)
            Workbook workbook = new Workbook("input.xlsx");

            // Select the target worksheet (first worksheet in this example)
            Worksheet worksheet = workbook.Worksheets[0];

            // Perform a shape operation: add a rectangle shape to the selected worksheet
            // Parameters: upper left row, upper left column, top, left, height, width
            Shape rectangle = worksheet.Shapes.AddRectangle(2, 2, 0, 0, 100, 50);
            rectangle.Name = "DemoRectangle";
            rectangle.FillFormat.ForeColor = System.Drawing.Color.LightBlue;

            // Save the modified workbook (uses Workbook.Save(string) method)
            workbook.Save("output.xlsx");
        }
    }
}
