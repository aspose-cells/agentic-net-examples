// Title: Refresh linked picture shapes after saving an Excel workbook with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells refresh linked picture shapes after workbook.save | How to update positions of linked images in Excel using Aspose.Cells .NET | Aspose.Cells Shape.Refresh method example for linked pictures | Iterate through worksheet shapes and call Refresh in Aspose.Cells C# | Handling unsupported shape types when refreshing linked shapes with Aspose.Cells
// Tags: Aspose.Cells linked picture shape refresh | refresh linked image positions after workbook save | C# iterate worksheet shapes Aspose.Cells | Aspose.Cells Shape.Refresh method | unsupported shape handling Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an existing Excel file, optionally saves a copy, loops through each worksheet and its shapes, attempts to invoke the dynamic Refresh method on linked picture shapes, logs any shapes that cannot be refreshed, and finally saves the workbook with refreshed linked images.
class RefreshLinkedShapesExample
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Optional: save the initial state
            workbook.Save("output_initial.xlsx");

            // Iterate through worksheets and their shapes
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (Shape shape in sheet.Shapes)
                {
                    try
                    {
                        // Attempt to refresh the shape (works for linked pictures)
                        dynamic linkedShape = shape;
                        linkedShape.Refresh();
                    }
                    catch (Exception ex)
                    {
                        // Refresh may not be supported for this shape; log and continue
                        Console.WriteLine($"Unable to refresh shape on sheet '{sheet.Name}': {ex.Message}");
                    }
                }
            }

            // Save the workbook after attempting to refresh linked shapes
            workbook.Save("output_refreshed.xlsx");
        }
        catch (Exception ex)
        {
            // General exception handling for unexpected errors
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
