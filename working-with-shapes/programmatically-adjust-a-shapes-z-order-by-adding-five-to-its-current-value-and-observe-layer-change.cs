// Title: How to increase a worksheet shape's Z-order by five using Aspose.Cells for .NET
// AI Prompts: Retrieve the first shape on a worksheet, add 5 to its ZOrderPosition, and save the workbook with Aspose.Cells in C#. | Read a shape's current Z-order, apply an offset of five layers, and persist the change to a new Excel file using the Aspose.Cells API. | Programmatically move an Excel shape forward five layers by updating its ZOrderPosition property in a .NET application.
// Common Searches: Aspose.Cells C# increase shape Z-order by specific number | set ZOrderPosition of an Excel shape using Aspose.Cells .NET | move worksheet shape forward multiple layers Aspose.Cells example | how to change layer order of shapes in an Excel file with Aspose.Cells | C# code to adjust Z-order of a shape in an existing workbook
// Tags: adjust shape Z-order Aspose.Cells | increase shape layer position .NET | modify ZOrderPosition Excel workbook | move Excel shape forward Aspose.Cells | shape ordering manipulation C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an existing Excel workbook, accesses the first worksheet, verifies that a shape exists, reads the shape's current ZOrderPosition, adds five to move it forward in the layer stack, outputs the change, and saves the modified workbook to a new file.
class Program
{
    static void Main()
    {
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        try
        {
            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure there is at least one shape
            if (worksheet.Shapes.Count == 0)
            {
                Console.WriteLine("No shapes found in the worksheet.");
                return;
            }

            // Retrieve the first shape
            Shape shape = worksheet.Shapes[0];

            // Get current Z-order position
            int currentZOrder = shape.ZOrderPosition;

            // Increase Z-order by five
            shape.ZOrderPosition = currentZOrder + 5;

            // Output the change
            Console.WriteLine($"Shape Z-order changed from {currentZOrder} to {shape.ZOrderPosition}");

            // Save the modified workbook
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
