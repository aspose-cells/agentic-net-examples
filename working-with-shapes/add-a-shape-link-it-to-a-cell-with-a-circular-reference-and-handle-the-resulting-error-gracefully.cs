// Title: Add a rectangle shape linked to a circular‑reference cell and gracefully handle the calculation error using Aspose.Cells for .NET
// AI Prompts: Insert a rectangle shape on the first worksheet, associate it with cell A1 that contains a circular reference, and wrap the CalculateFormula call in a try‑catch to capture any exception. | If the formula calculation fails, write a custom error message into the circular‑reference cell and then save the workbook to a chosen file path.
// Common Searches: Aspose.Cells how to catch circular reference exception during workbook calculation | C# add rectangle shape to worksheet and link it to a cell using Aspose.Cells | FreeFloating shape placement parameters in Aspose.Cells .NET example | Replace circular reference formula with error text after calculation failure Aspose.Cells | Save Excel file after handling formula errors with Aspose.Cells
// Tags: add rectangle shape worksheet Aspose.Cells | circular reference formula error handling Aspose.Cells | freefloating shape placement Aspose.Cells .NET | link shape to cell Aspose.Cells | save workbook after formula exception Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new workbook, sets a circular reference formula in cell A1, adds a rectangle shape positioned at row 1 column 0, attempts to calculate all formulas, catches the resulting exception, writes a custom error message into the cell, and saves the workbook as CircularReferenceDemo.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Set a circular reference formula in cell A1 (e.g., =A1+1)
            Cell circularCell = sheet.Cells["A1"];
            circularCell.Formula = "=A1+1";

            // Add a rectangle shape to the worksheet
            // Parameters: shape type, upper left row, upper left column, top, left, height, width
            Shape shape = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle, // shape type
                1,   // upper left row
                0,   // upper left column
                0,   // top (in points)
                100, // left (in points)
                50,  // height (in points)
                100  // width (in points)
            );

            // Configure shape properties
            shape.Name = "CircularReferenceShape";
            shape.Placement = PlacementType.FreeFloating;

            // Note: SetPosition is not available in this version of Aspose.Cells.
            // The shape is already positioned by the AddShape parameters.

            // Attempt to calculate formulas and handle any errors gracefully
            try
            {
                workbook.CalculateFormula();
            }
            catch (Exception ex)
            {
                // Aspose.Cells may throw a generic exception for circular references
                Console.WriteLine("Formula calculation error: " + ex.Message);
                circularCell.PutValue("Error: Circular Ref");
            }

            // Save the workbook to a file
            string outputPath = "CircularReferenceDemo.xlsx";
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine("Workbook saved successfully to " + Path.GetFullPath(outputPath));
            }
            catch (Exception ex)
            {
                Console.WriteLine("Failed to save workbook: " + ex.Message);
            }
        }
        catch (Exception e)
        {
            // General exception handling for unexpected errors
            Console.WriteLine("An error occurred: " + e.Message);
        }
    }
}
