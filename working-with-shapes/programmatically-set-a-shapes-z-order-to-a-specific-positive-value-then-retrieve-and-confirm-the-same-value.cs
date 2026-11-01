// Title: How to set and read a shape's Z-order position in an Aspose.Cells workbook using C#
// AI Prompts: Create a new workbook, add a rectangle shape to the first worksheet, set its ZOrderPosition to 5, then output the Z-order value with Aspose.Cells for .NET. | Show how to change a shape's stacking order and retrieve the ZOrderPosition property in a C# Aspose.Cells example.
// Common Searches: Aspose.Cells C# set shape ZOrderPosition to specific value | Retrieve shape stacking order after setting ZOrderPosition in Aspose.Cells .NET | C# example for changing and confirming shape Z-order in Excel file using Aspose.Cells | How to programmatically control shape order in an Aspose.Cells workbook
// Tags: Aspose.Cells shape ZOrderPosition set C# | Aspose.Cells retrieve shape Z-order .NET | Aspose.Cells add rectangle shape worksheet C# | Aspose.Cells workbook shape stacking order example | Aspose.Cells Z-order verification C#

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a new workbook, inserts a rectangle shape, assigns its ZOrderPosition property to a positive integer (5), reads the property back, prints the value, and saves the file as ZOrderDemo.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet
            // Parameters: shape type, upper left row, upper left column, top, left, height, width
            Shape shape = worksheet.Shapes.AddShape(
                MsoDrawingType.Rectangle, // shape type
                1,   // upper left row
                0,   // upper left column
                0,   // top offset (pixels)
                0,   // left offset (pixels)
                100, // height (pixels)
                100  // width (pixels)
            );

            // Set the Z-order to a specific positive value (e.g., 5)
            shape.ZOrderPosition = 5;

            // Retrieve the Z-order to confirm it was set correctly
            int currentZOrder = shape.ZOrderPosition;

            // Output the result
            Console.WriteLine("Shape Z-order is set to: " + currentZOrder);

            // Save the workbook (optional, just to demonstrate full lifecycle)
            workbook.Save("ZOrderDemo.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
