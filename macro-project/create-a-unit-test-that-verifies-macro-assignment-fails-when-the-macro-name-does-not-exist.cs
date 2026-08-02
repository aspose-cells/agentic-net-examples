// Title: Unit test to confirm shape macro assignment fails when the macro is missing in Aspose.Cells for .NET
// Description: Creates a macro‑enabled workbook, adds a rectangle shape, sets its MacroName to a non‑existent VBA procedure, and saves as Xlsm. The test asserts that Aspose.Cells throws a CellsException, proving that invalid macro references are detected during save.
// Keywords: Aspose.Cells unit test | C# macro validation | shape MacroName exception | CellsException on save | macro‑enabled workbook testing | .NET Aspose.Cells macro error | invalid VBA procedure detection
// Common Searches: Aspose.Cells test for missing macro on shape | how to assert CellsException for invalid MacroName | unit test macro validation Aspose.Cells .NET | shape macro assignment error handling | verify macro existence before saving workbook
// Developer Intent: Validate that assigning a non‑existent macro to a shape triggers a CellsException when the workbook is saved as a macro‑enabled file.
// Use Cases: Automated CI checks for broken macro references in generated Excel files | Quality‑gate testing in document‑generation pipelines that use Aspose.Cells | Ensuring reliable workbook saves by catching invalid VBA procedure names early
// AI Prompts: Generate an MSTest method that creates a workbook with macros enabled, adds a rectangle shape, assigns a non‑existent MacroName, saves as Xlsm, and asserts that a CellsException is thrown. | Write a NUnit test for Aspose.Cells that verifies saving a macro‑enabled workbook fails when a shape references a missing VBA macro. | Provide an xUnit example that checks the exception message returned by Aspose.Cells when an invalid MacroName is set on a shape.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsTests
{
    // Creates a macro‑enabled workbook, adds a rectangle shape, sets its MacroName to a non‑existent VBA procedure, and saves as Xlsm. The test asserts that Aspose.Cells throws a CellsException, proving that invalid macro references are detected during save.
    class Program
    {
        static void Main(string[] args)
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

                // Assign a macro name that does not exist in the workbook
                shape.MacroName = "NonExistingMacro()";

                // Attempt to save as a macro‑enabled file; this should throw a CellsException
                string outputPath = "NonExistingMacroTest.xlsm";

                // Ensure any existing file is deleted to avoid FileNotFoundException on overwrite
                if (File.Exists(outputPath))
                {
                    File.Delete(outputPath);
                }

                try
                {
                    workbook.Save(outputPath, SaveFormat.Xlsm);
                    Console.WriteLine("Test Failed: Expected exception was not thrown.");
                }
                catch (CellsException ex)
                {
                    // Expected path – the macro cannot be resolved
                    Console.WriteLine($"Test Passed: Caught expected CellsException -> {ex.Message}");
                }
                catch (Exception ex)
                {
                    // Any other exception indicates an unexpected failure
                    Console.WriteLine($"Test Failed: Unexpected exception -> {ex.GetType().Name}: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                // General safety net for any unforeseen errors
                Console.WriteLine($"Unhandled exception: {ex.GetType().Name}: {ex.Message}");
            }
        }
    }
}
