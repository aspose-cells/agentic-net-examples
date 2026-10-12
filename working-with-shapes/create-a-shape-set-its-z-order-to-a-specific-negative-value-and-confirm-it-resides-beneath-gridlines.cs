// Title: Create a rectangle shape in Aspose.Cells for .NET, set a negative Z-order to position it behind gridlines, and verify placement
// AI Prompts: Add a rectangle shape to a worksheet and assign its ZOrderPosition a negative integer so the shape appears beneath the gridlines. | Write code that checks the ZOrderPosition of a shape and outputs a message indicating whether it is positioned behind the gridlines. | Save the workbook to an .xlsx file after positioning the shape and ensure the file is created without errors.
// Common Searches: Aspose.Cells C# set shape ZOrderPosition negative to hide behind gridlines | how to place a rectangle shape under worksheet gridlines using Aspose.Cells for .NET | verify shape Z-order value in Aspose.Cells example code | C# Aspose.Cells add shape and change Z-order behind gridlines
// Tags: add rectangle shape Aspose.Cells C# | set negative ZOrderPosition Aspose.Cells | shape behind gridlines Aspose.Cells | verify shape Z-order Aspose.Cells | save workbook as xlsx Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates adding a rectangle shape to a worksheet, assigning a negative ZOrderPosition to place the shape behind gridlines, checking the Z-order value, and saving the workbook as ShapeBehindGridlines.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet
            // Parameters: shape type, upper left row, upper left column, upper left offset (pixels),
            //             upper left offset (pixels), height (pixels), width (pixels)
            Shape shape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 2, 2, 0, 0, 100, 50);

            // Set the Z-order to a specific negative value so it is placed behind gridlines
            shape.ZOrderPosition = -5;

            // Confirm the shape's Z-order is negative (i.e., beneath gridlines)
            if (shape.ZOrderPosition < 0)
            {
                Console.WriteLine($"Shape Z-order is {shape.ZOrderPosition}, which places it beneath the gridlines.");
            }
            else
            {
                Console.WriteLine($"Shape Z-order is {shape.ZOrderPosition}, which does NOT place it beneath the gridlines.");
            }

            // Save the workbook (lifecycle rule: save)
            workbook.Save("ShapeBehindGridlines.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
