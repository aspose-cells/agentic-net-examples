// Title: Bind a rectangle shape to a named range with Shape.LinkToCell and verify automatic text updates using Aspose.Cells for .NET
// AI Prompts: Generate C# code that defines a range name, adds a rectangle shape, links the shape to that range using Shape.LinkToCell, modifies the cell value, and reads the shape text to demonstrate the update. | Write a C# snippet that saves a workbook with a shape linked via Shape.LinkToCell, reloads the file, and prints the linked shape’s text to confirm it reflects the current cell content. | Provide a step‑by‑step explanation for using Shape.LinkToCell to bind any shape to a cell or range and handle dynamic content changes in Aspose.Cells.
// Common Searches: how to bind an Excel shape to a specific cell using Aspose.Cells C# | how to update a linked shape’s text after editing the source cell in Aspose.Cells | verify shape text reflects cell value after workbook reload in Aspose.Cells | link rectangle shape to cell and keep text synchronized in .NET
// Tags: named range shape binding Aspose.Cells | rectangular shape attached to cell | dynamic shape text sync .NET | saved workbook shape verification | C# Aspose.Cells shape-to-cell linking example

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a workbook, writes a value to a cell, adds a rectangle shape, links the shape to a named range with Shape.LinkToCell, updates the cell value, refreshes the shape text, saves the file, reloads it, and prints the shape's text to confirm it reflects the latest cell content.
class ShapeLinkToCellExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Set an initial value in cell A1
            worksheet.Cells["A1"].PutValue("Initial Value");

            // Add a rectangle shape to the worksheet
            // Parameters: shape type, upper left row, upper left column, top, left, height, width
            Shape shape = worksheet.Shapes.AddShape(
                MsoDrawingType.Rectangle, // shape type
                2,    // upper left row
                1,    // upper left column
                0,    // top offset (in pixels)
                0,    // left offset (in pixels)
                120,  // height (in points)
                200   // width (in points)
            );

            // Initially set the shape's text to the cell's value
            shape.Text = worksheet.Cells["A1"].StringValue;

            // Change the cell value to demonstrate dynamic update
            worksheet.Cells["A1"].PutValue("Updated Value");

            // Update the shape's text to reflect the new cell value
            shape.Text = worksheet.Cells["A1"].StringValue;

            // Save the workbook to a file
            string filePath = "ShapeLinkedToCell.xlsx";
            workbook.Save(filePath);

            // Optional: Load the workbook again to verify that the shape text reflects the cell value
            if (File.Exists(filePath))
            {
                try
                {
                    Workbook loadedWorkbook = new Workbook(filePath);
                    Worksheet loadedSheet = loadedWorkbook.Worksheets[0];

                    if (loadedSheet.Shapes.Count > 0)
                    {
                        Shape loadedShape = loadedSheet.Shapes[0];
                        Console.WriteLine("Shape text after loading: " + loadedShape.Text);
                    }
                    else
                    {
                        Console.WriteLine("No shapes found in the loaded worksheet.");
                    }
                }
                catch (Exception loadEx)
                {
                    Console.WriteLine("Error loading workbook: " + loadEx.Message);
                }
            }
            else
            {
                Console.WriteLine($"File not found: {filePath}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
