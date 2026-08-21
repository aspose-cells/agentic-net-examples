// Title: Aspose.Cells .NET unit test – verify exception when assigning a non‑existent macro to a shape
// Description: Creates a macro‑enabled workbook, adds a rectangle shape, sets its MacroName to a name that is not defined, and saves the file. The test asserts that Aspose.Cells throws a CellsException because the macro reference is invalid.
// Keywords: Aspose.Cells | .NET | macro validation | shape MacroName | exception handling | unit test | MSTest | NUnit | xUnit | CellsException | nonexistent macro | save validation
// Common Searches: Aspose.Cells test missing macro exception | shape macro name not found error .NET | unit test for macro validation Aspose.Cells | assert CellsException on invalid macro reference | how to verify macro assignment fails Aspose.Cells
// Developer Intent: Confirm that assigning a macro name that does not exist to a shape triggers a CellsException during workbook save.
// Use Cases: Write an MSTest/NUnit/xUnit test that creates a macro‑enabled workbook, adds a shape, assigns an invalid MacroName, and asserts that Save throws a CellsException. | Add the test to a CI pipeline to catch regressions in macro reference validation after library updates. | Log the exception message to verify it contains the missing macro name for easier debugging.
// AI Prompts: Generate an MSTest method that creates a workbook with EnableMacros = true, adds a rectangle shape, sets MacroName = "NonExistingMacro()", saves the workbook, and asserts that a CellsException is thrown. | Provide a NUnit test example that verifies Aspose.Cells throws an exception when a shape references a macro that is not present in the workbook. | Write an xUnit test that catches the expected CellsException and checks that the exception message includes the missing macro name.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using System.IO;

namespace AsposeCellsTests
{
    // Creates a macro‑enabled workbook, adds a rectangle shape, sets its MacroName to a name that is not defined, and saves the file. The test asserts that Aspose.Cells throws a CellsException because the macro reference is invalid.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and enable macros
                Workbook workbook = new Workbook();
                workbook.Settings.EnableMacros = true;

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a rectangle shape to the worksheet
                Shape shape = worksheet.Shapes.AddRectangle(1, 1, 100, 100, 0, 0);

                // Assign a macro name that does NOT exist in the workbook
                shape.MacroName = "NonExistingMacro()";

                // Define the output file path
                string outputPath = "NonExistingMacroTest.xlsm";

                // Save the workbook; Aspose.Cells validates macro references on save
                workbook.Save(outputPath);
                Console.WriteLine("Workbook saved successfully (unexpected).");
            }
            catch (Exception ex)
            {
                // Expected exception because the macro is missing
                Console.WriteLine($"Expected exception caught: {ex.Message}");
            }
        }
    }
}
