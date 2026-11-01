// Title: Read the shadow color of a worksheet shape and log its RGB components using Aspose.Cells for .NET
// AI Prompts: Write C# code that opens an Excel workbook with Aspose.Cells, accesses the first shape on a worksheet, retrieves the ShadowEffect color, and prints the R, G, B values to the console. | Show how to safely extract a shape's shadow color in Aspose.Cells, handling cases where the shadow effect or its color is missing, and output the RGB components for logging.
// Common Searches: Aspose.Cells C# get RGB values of a shape's shadow color | how to read shadow effect color from an Excel shape using Aspose.Cells | log shape shadow color components in a .NET workbook | C# retrieve shadow color of the first worksheet shape with Aspose.Cells | extract shape shadow color when it may be null Aspose.Cells example
// Tags: Aspose.Cells shape shadow color extraction | C# read shape shadow RGB Aspose.Cells | Aspose.Cells shadow effect color handling | Excel worksheet shape shadow logging .NET | Aspose.Cells retrieve shape shadow properties

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an Excel workbook, checks for shapes on the first worksheet, accesses the first shape's ShadowEffect, safely obtains its color (handling nulls), converts it to System.Drawing.Color, and writes the R, G, B components to the console. It includes error handling for missing files, absent shapes, or missing shadow effects.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (or specify the desired one)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure there is at least one shape on the worksheet
            if (worksheet.Shapes.Count == 0)
            {
                Console.WriteLine("No shapes found on the worksheet.");
                return;
            }

            // Retrieve the first shape (adjust index as needed)
            Shape shape = worksheet.Shapes[0];

            // Access the shadow effect; it may be null if no shadow is set
            ShadowEffect shadowEffect = shape.ShadowEffect;
            if (shadowEffect == null)
            {
                Console.WriteLine("The shape does not have a shadow effect.");
                return;
            }

            // Convert Aspose.Cells.CellsColor to System.Drawing.Color
            Color shadowColor = shadowEffect.Color?.Color ?? Color.Empty;

            // Log the RGB components of the shadow color
            Console.WriteLine($"Shadow Color RGB: R={shadowColor.R}, G={shadowColor.G}, B={shadowColor.B}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
