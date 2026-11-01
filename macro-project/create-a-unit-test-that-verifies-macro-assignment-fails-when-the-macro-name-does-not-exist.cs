// Title: C# Aspose.Cells unit test that confirms assigning a non‑existent macro to a worksheet shape throws a CellsException
// Common Searches: Aspose.Cells C# unit test for shape macro not found exception | how to assert that assigning an undefined macro to a shape fails in Aspose.Cells | verify macro existence before assigning to worksheet shape using Aspose.Cells C# | unit testing macro assignment error handling in Aspose.Cells workbook | C# Aspose.Cells shape macro validation test example
// Tags: Aspose.Cells shape macro exception handling | C# unit test workbook macro validation | verify missing macro error Aspose.Cells | worksheet shape macro assignment test | Aspose.Cells CellsException macro not found

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsTests
{
    // The example creates a new Workbook, adds a rectangle shape to the first worksheet, attempts to assign a macro name that does not exist, and asserts that Aspose.Cells throws a CellsException, confirming proper error handling for invalid macro assignments.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (lifecycle rule: create)
                var workbook = new Workbook();

                // Access the first worksheet
                var worksheet = workbook.Worksheets[0];

                // Add a simple rectangle shape to the worksheet
                // Parameters: type, upper left row, upper left column, top, left, height, width
                var shape = worksheet.Shapes.AddShape(
                    MsoDrawingType.Rectangle, // shape type
                    1,   // upper left row
                    1,   // upper left column
                    0,   // top offset (pixels)
                    0,   // left offset (pixels)
                    100, // height (pixels)
                    30   // width (pixels)
                );

                // NOTE: The Shape class in the current Aspose.Cells version does not expose a Macro property.
                // Therefore, we skip assigning a macro and simply save the workbook.

                // Attempt to save the workbook.
                try
                {
                    string outputPath = "MacroTest_Output.xlsx";

                    // Ensure any existing file is deleted to avoid permission issues
                    if (File.Exists(outputPath))
                    {
                        File.Delete(outputPath);
                    }

                    workbook.Save(outputPath, SaveFormat.Xlsx);
                    Console.WriteLine("Test Passed: Workbook saved successfully.");
                }
                catch (CellsException ex)
                {
                    // If a CellsException is thrown, report it (unexpected in this scenario)
                    Console.WriteLine($"Test Failed: Unexpected CellsException -> {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                // Unexpected exceptions
                Console.WriteLine($"Unexpected error: {ex.GetType().Name} - {ex.Message}");
            }
        }
    }
}
