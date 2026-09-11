// Title: How to set a shape's Z-order just above another shape in an Excel worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Locate the shape called 'MyShape' and assign its Z-order to be one level higher than the shape 'ReferenceShape' via the Aspose.Cells C# library. | Raise a target shape above a reference shape in an Excel worksheet by updating the shape's stacking order with Aspose.Cells. | Reorder two worksheet shapes so the first appears directly on top of the second using Aspose.Cells in a .NET application.
// Common Searches: Aspose.Cells C# bring one shape to front of another in Excel | how to change shape Z-order in an Excel file using Aspose.Cells | set shape stacking order relative to another shape with Aspose.Cells .NET | C# code to move Excel shape above another shape programmatically | adjust worksheet shape order using Aspose.Cells API
// Tags: Aspose.Cells modify shape order C# | Excel worksheet shape hierarchy Aspose.Cells | C# adjust shape ZOrderPosition Aspose.Cells | Aspose.Cells reorder worksheet shapes .NET | set shape front/back order Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// The example loads an existing workbook, retrieves two shapes by name, reads the Z-order of the reference shape, sets the target shape's ZOrderPosition to one higher, and saves the workbook with the updated stacking order.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index or name as needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Get the collection of shapes on the worksheet
            ShapeCollection shapes = sheet.Shapes;

            // Retrieve the shape to move (replace with actual name or index)
            Shape shapeToMove = shapes["MyShape"];
            if (shapeToMove == null)
            {
                Console.WriteLine("Error: Shape \"MyShape\" not found.");
                return;
            }

            // Retrieve the reference shape that should be directly below the moved shape
            Shape referenceShape = shapes["ReferenceShape"];
            if (referenceShape == null)
            {
                Console.WriteLine("Error: Shape \"ReferenceShape\" not found.");
                return;
            }

            // Get the Z-order position of the reference shape
            int referenceZOrder = referenceShape.ZOrderPosition;

            // Set the Z-order of the target shape to be just above the reference shape
            shapeToMove.ZOrderPosition = referenceZOrder + 1;

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors during processing
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
