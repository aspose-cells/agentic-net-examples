// Title: Link a rectangle shape’s text to cell K3 for automatic updates with Aspose.Cells for .NET (C#)
// AI Prompts: Create a rectangle shape on a worksheet and bind its Text property to the formula "=K3" so the shape displays the cell’s current value. | Configure the shape’s TextHorizontalAlignment to Center to align the linked text inside the rectangle. | Save the workbook after setting up the shape‑cell link to produce an Excel file that reflects any change in K3. | Wrap the shape‑linking code in a try‑catch block to handle potential Aspose.Cells exceptions.
// Common Searches: Aspose.Cells C# how to bind shape text to a cell value | link rectangle shape text to Excel cell K3 using Aspose.Cells for .NET | automatic shape text update when cell changes Aspose.Cells example | set shape text formula =K3 in Aspose.Cells C# | center align text inside a linked shape Aspose.Cells
// Tags: shape text formula binding Aspose.Cells | rectangle shape cell link C# | horizontal text centering for shape Aspose.Cells | auto updating shape from cell value .NET | save workbook with linked shape Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The program creates a new workbook, adds a rectangle shape to the first worksheet, links the shape's text to cell K3 using the "=K3" formula, centers the text horizontally, and saves the file as LinkedShape.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a rectangle shape (rows/columns + size in points)
            // In newer Aspose.Cells versions AddShape returns a Shape object directly
            Shape shape = worksheet.Shapes.AddShape(
                MsoDrawingType.Rectangle,
                2, // upper left row
                2, // upper left column
                5, // lower right row
                5, // lower right column
                100, // height (points)
                200  // width (points)
            );

            // Link the shape's text to cell K3
            shape.Text = "=K3";

            // Center align text inside the shape
            shape.TextHorizontalAlignment = TextAlignmentType.Center;

            // Save the workbook
            workbook.Save("LinkedShape.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
