// Title: Read the shadow color of a worksheet shape and log its RGB values using Aspose.Cells for .NET
// AI Prompts: Retrieve the ShadowEffect.Color of a shape in an Excel worksheet and output the R, G, B components with Aspose.Cells in C#. | Write C# code that checks whether a shape has a shadow, extracts its color, and prints the RGB values to the console using Aspose.Cells. | Generate a snippet that loads an .xlsx file, accesses the first shape, and logs the shadow color's RGB channels via the Aspose.Cells API.
// Common Searches: how to get shape shadow color in Aspose.Cells C# | Aspose.Cells read RGB values of shape shadow effect | C# Aspose.Cells retrieve shadow color from Excel shape | log shadow effect color of a shape using Aspose.Cells .NET | extract shape shadow color from worksheet with Aspose.Cells
// Tags: Aspose.Cells read shape shadow color | C# extract shape shadow RGB | Aspose.Cells shadow effect color logging | Excel shape shadow property Aspose.Cells | retrieve shape shadow color .NET

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.Drawing;
using System.IO;

// The example loads an Excel workbook, accesses the first worksheet, obtains the first shape, checks for a ShadowEffect, converts its CellsColor to System.Drawing.Color, and writes the shadow's R, G, B values to the console.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the input workbook
            string filePath = "input.xlsx";

            // Verify that the file exists to avoid FileNotFoundException
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                return;
            }

            // Load the workbook from the file
            Workbook workbook = new Workbook(filePath);

            // Access the first worksheet (adjust index as needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one shape
            if (sheet.Shapes.Count > 0)
            {
                // Retrieve the first shape (or use sheet.Shapes["ShapeName"] to get by name)
                Shape shape = sheet.Shapes[0];

                // Get the shadow effect of the shape (may be null if no shadow is set)
                ShadowEffect shadow = shape.ShadowEffect;

                if (shadow != null)
                {
                    // Convert Aspose.Cells.CellsColor to System.Drawing.Color
                    Color shadowColor = shadow.Color.Color; // CellsColor.Color returns System.Drawing.Color

                    // Log the RGB components of the shadow color
                    Console.WriteLine($"Shadow Color RGB: R={shadowColor.R}, G={shadowColor.G}, B={shadowColor.B}");
                }
                else
                {
                    Console.WriteLine("The shape does not have a shadow effect.");
                }
            }
            else
            {
                Console.WriteLine("No shapes found in the worksheet.");
            }
        }
        catch (Exception ex)
        {
            // Handle unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
