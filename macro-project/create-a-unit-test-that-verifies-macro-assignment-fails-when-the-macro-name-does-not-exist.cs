using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace ShapeMacroDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                AssignNonExistingMacro_ShouldFail();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }

        // Test: assigning a macro that does not exist should cause an exception on save.
        static void AssignNonExistingMacro_ShouldFail()
        {
            try
            {
                // Create a new workbook (lifecycle: create)
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a rectangle shape to the worksheet
                Shape shape = worksheet.Shapes.AddRectangle(1, 1, 100, 100, 0, 0);

                // Assign a non‑existing macro name
                shape.MacroName = "NonExistingMacro()";

                // Save the workbook to a memory stream (lifecycle: save)
                using (MemoryStream ms = new MemoryStream())
                {
                    workbook.Save(ms, SaveFormat.Xlsm);
                }

                // If we reach this point, no exception was thrown – the test failed.
                Console.WriteLine("Test Failed: No exception was thrown when saving with a non‑existing macro.");
            }
            catch (Exception ex)
            {
                // Expected path: an exception should be thrown because the macro cannot be resolved.
                Console.WriteLine($"Test Passed: Caught expected exception – {ex.Message}");
            }
        }
    }
}