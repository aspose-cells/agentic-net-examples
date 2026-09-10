// Title: Retrieve the Shapes collection from the active worksheet with Aspose.Cells for .NET (C#)
// AI Prompts: Use Aspose.Cells to obtain the Shapes collection of the workbook's currently active worksheet. | Iterate over each Shape in the active sheet and output its Name and Type using C#. | Save the workbook after processing the shape collection if any changes are required.
// Common Searches: C# Aspose.Cells get shapes from the selected worksheet | How to list all drawing objects on the active sheet using Aspose.Cells | Retrieve shape names and types from the current worksheet in Aspose.Cells .NET | Example code for enumerating worksheet shapes with Aspose.Cells C#
// Tags: retrieve shapes collection Aspose.Cells .NET | enumerate worksheet drawing objects C# | active sheet shape enumeration Aspose.Cells | access shape properties worksheet Aspose.Cells | process worksheet shapes C# example

using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads a workbook, accesses the active worksheet, retrieves its Shapes collection, iterates through each shape to display its name and type, and optionally saves the workbook after any modifications.
class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Get the currently selected (active) worksheet
        Worksheet worksheet = workbook.Worksheets[workbook.Worksheets.ActiveSheetIndex];

        // Retrieve the Shapes collection of the selected worksheet
        ShapeCollection shapes = worksheet.Shapes;

        // Example processing: iterate through each shape
        foreach (Shape shape in shapes)
        {
            // Output basic information about the shape
            System.Console.WriteLine($"Shape Name: {shape.Name}, Type: {shape.Type}");
        }

        // Save the workbook if any modifications were made (optional)
        workbook.Save("output.xlsx");
    }
}
