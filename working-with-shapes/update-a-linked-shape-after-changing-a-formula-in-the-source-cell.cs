// Title: Update a linked shape after modifying a cell formula using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# recalculate formulas and refresh linked shapes in an existing workbook | How to change a cell formula and automatically update a linked picture with Aspose.Cells | Retrieve a named shape after updating formulas using the Aspose.Cells .NET API | Refresh Excel shapes linked to cells after a formula change with Aspose.Cells CalculateFormula | C# code to modify A1 formula and update a linked chart shape in Aspose.Cells
// Tags: calculateformula linked shape update | cell formula change triggers shape refresh | worksheet shape retrieval by name | auto update linked picture after formula edit | excel workbook recalculate shapes Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an existing workbook, changes the formula in cell A1, calls CalculateFormula to recalculate all formulas and automatically refresh any shapes linked to cells, optionally retrieves a shape named "MyLinkedShape" for further manipulation, and then saves the updated workbook.
class UpdateLinkedShapeExample
{
    static void Main()
    {
        const string sourcePath = "SourceWorkbook.xlsx";
        const string outputPath = "UpdatedWorkbook.xlsx";

        try
        {
            // Verify that the source workbook exists to avoid FileNotFoundException.
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Error: The file \"{sourcePath}\" was not found.");
                return;
            }

            // Load the workbook.
            Workbook workbook = new Workbook(sourcePath);

            // Access the first worksheet.
            Worksheet worksheet = workbook.Worksheets[0];

            // Change the formula in the source cell (e.g., A1).
            Cell sourceCell = worksheet.Cells["A1"];
            sourceCell.Formula = "=SUM(B1:B10)";

            // Recalculate all formulas in the workbook.
            // This also refreshes any shapes that are linked to cells.
            workbook.CalculateFormula();

            // Optionally, access a specific shape by name.
            // Note: Aspose.Cells automatically updates linked shapes after CalculateFormula,
            // so no explicit update method is required.
            try
            {
                Shape linkedShape = worksheet.Shapes["MyLinkedShape"];
                if (linkedShape != null)
                {
                    Console.WriteLine($"Shape \"{linkedShape.Name}\" was found.");
                    // Additional shape manipulation can be performed here if needed.
                }
                else
                {
                    Console.WriteLine("Shape \"MyLinkedShape\" was not found.");
                }
            }
            catch (Exception shapeEx)
            {
                Console.WriteLine($"An error occurred while handling the shape: {shapeEx.Message}");
            }

            // Save the updated workbook.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message.
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
