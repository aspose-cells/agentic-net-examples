// Title: Programmatically set and verify a shape's Z‑order position in an Excel worksheet with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells how to assign a specific Z-order to a worksheet shape | retrieve ZOrderPosition of a shape after setting it in Aspose.Cells .NET | example code for changing shape stacking order in an Excel file using Aspose.Cells | Aspose.Cells set shape Z-order value and confirm it programmatically | save workbook after modifying shape Z-order with Aspose.Cells C#
// Tags: Aspose.Cells shape ZOrderPosition property | set shape stacking order .NET | retrieve shape Z-order Aspose.Cells | add rectangle shape to worksheet Aspose.Cells | save workbook with modified shape order Aspose.Cells | Excel shape Z-order manipulation C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsZOrderDemo
{
    // The sample creates a new Workbook, adds a rectangle shape to the first worksheet, sets its ZOrderPosition to a specific value (5), reads the property back to confirm the assignment, writes the result to the console, and saves the workbook as ZOrderDemo.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Add a rectangle shape to the worksheet
                // Parameters: shape type, upper left row, upper left column, top, left, height, width
                Shape shape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 2, 2, 0, 0, 100, 100);

                // Set the shape's Z-order position (available as ZOrderPosition)
                shape.ZOrderPosition = 5;

                // Retrieve the Z-order position
                int currentZOrder = shape.ZOrderPosition;

                // Output the Z-order to confirm it was set correctly
                Console.WriteLine("Shape Z-order: " + currentZOrder);
                if (currentZOrder == 5)
                {
                    Console.WriteLine("Z-order set correctly.");
                }
                else
                {
                    Console.WriteLine("Z-order mismatch.");
                }

                // Define output file path
                string outputPath = "ZOrderDemo.xlsx";

                // Save the workbook (optional, demonstrates lifecycle usage)
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
