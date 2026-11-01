// Title: How to group three worksheet shapes into a named composite object with Aspose.Cells for .NET (C#)
// AI Prompts: Use Aspose.Cells in C# to retrieve shapes named "Shape 1", "Shape 2", and "Shape 3" from a worksheet, group them into a single composite shape, set its Name to "MyCompositeObject", and save the workbook. | Create a named shape group in a .NET workbook by invoking Worksheet.Shapes.Group on selected shapes and assigning a custom Name property with Aspose.Cells.
// Common Searches: how to combine several worksheet drawings into one object using Aspose.Cells C# | setting a custom identifier for a shape group in an Excel file with Aspose.Cells | retrieving shapes by name and creating a composite shape in .NET | saving an Excel workbook after grouping drawings via Aspose.Cells API | example of Worksheet.Shapes.Group method in C#
// Tags: worksheet shapes grouping Aspose.Cells | named composite shape .NET | Aspose.Cells Shape.Group method | C# Excel shape group example | custom shape name Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The code loads (or creates) a workbook, fetches three shapes named "Shape 1", "Shape 2", and "Shape 3" from the first worksheet, groups them into a composite shape, assigns the name "MyCompositeObject" to the group, and saves the workbook as "output.xlsx".
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Load the workbook; create a new one if the input file does not exist.
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                Console.WriteLine($"Input file '{inputPath}' not found. Creating a new workbook.");
                workbook = new Workbook(); // creates a default workbook with one worksheet
            }

            Worksheet worksheet = workbook.Worksheets[0];

            // Retrieve shapes by name. Ensure the shapes exist before grouping.
            Shape shape1 = worksheet.Shapes["Shape 1"];
            Shape shape2 = worksheet.Shapes["Shape 2"];
            Shape shape3 = worksheet.Shapes["Shape 3"];

            // Group the selected shapes into a composite object.
            var composite = worksheet.Shapes.Group(new Shape[] { shape1, shape2, shape3 });

            // Assign a descriptive name to the composite shape for later reference.
            composite.Name = "MyCompositeObject";

            // Save the workbook with the new grouped shape.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
