// Title: Log worksheet, shape name and assigned macro from an Excel workbook using Aspose.Cells in C#
// Common Searches: how to enumerate all shapes in an Excel workbook and get their assigned macro using Aspose.Cells in C# | C# code to print worksheet name, shape name and macro stored in alternative text with Aspose.Cells | retrieve macro name from shape alternative text in an Excel file using Aspose.Cells .NET | list shape macro assignments for each worksheet in a workbook with Aspose.Cells library
// Tags: Aspose.Cells read shape macro name | log worksheet shape macro assignments C# | enumerate workbook shapes Aspose.Cells | extract alternative text macro Aspose.Cells | iterate worksheet shapes .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;   // Required for Shape and ShapeCollection

// The example loads an Excel workbook with Aspose.Cells, iterates through each worksheet and its ShapeCollection, reads the macro name from each shape's AlternativeText property, and writes the worksheet name, shape name, and macro name to the console.
class MacroAssignmentLogger
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Access the collection of shapes (e.g., buttons, checkboxes) on the worksheet
                ShapeCollection shapes = sheet.Shapes;

                // Loop through each shape to retrieve macro assignment details
                foreach (Shape shape in shapes)
                {
                    // Aspose.Cells older versions do not expose a direct Macro property.
                    // The macro name can be stored in the AlternativeText of the shape.
                    // Adjust this logic if your version provides a dedicated Macro property.
                    string macroName = shape.AlternativeText;

                    if (!string.IsNullOrEmpty(macroName))
                    {
                        Console.WriteLine($"Worksheet: {sheet.Name}, Shape Name: {shape.Name}, Macro: {macroName}");
                    }
                }
            }

            // If you made any changes and need to save the workbook, uncomment the line below
            // workbook.Save("output.xlsx");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
