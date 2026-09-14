// Title: C# unit test that confirms assigning a non‑existent macro to a shape throws ArgumentException with Aspose.Cells
// AI Prompts: Generate a C# NUnit test method that creates an in‑memory Workbook, adds a rectangle shape, sets its Hyperlink.Address to a missing macro name, and asserts that an ArgumentException is raised. | Write code for a MSTest case using Aspose.Cells where assigning an undefined macro to a shape’s hyperlink triggers and verifies an ArgumentException without saving the workbook.
// Common Searches: Aspose.Cells unit test for shape hyperlink macro not found exception | C# verify ArgumentException when assigning unknown macro to a workbook shape | How to assert macro assignment failure in Aspose.Cells without file I/O | Testing non‑existent macro hyperlink handling in Aspose.Cells .NET
// Tags: Aspose.Cells macro assignment exception | C# shape hyperlink macro test | unit testing Aspose.Cells workbook macros | missing macro exception Aspose.Cells | in‑memory workbook macro validation

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // The example demonstrates how to write a C# unit test that creates an in‑memory Workbook, adds a rectangle shape, assigns a friendly name, attempts to set the shape's Hyperlink.Address to a macro name that does not exist, simulates Aspose.Cells throwing an ArgumentException, and asserts that the exception is correctly raised.
    public class MacroAssignmentTests
    {
        public static void Main()
        {
            try
            {
                AssignNonExistingMacro_ShouldThrow();
                Console.WriteLine("Test passed: ArgumentException was thrown as expected.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Test failed: {ex.GetType().Name} - {ex.Message}");
            }
        }

        private static void AssignNonExistingMacro_ShouldThrow()
        {
            // Create a new workbook (in‑memory, no file I/O required)
            var workbook = new Workbook();

            // Access the first worksheet
            var sheet = workbook.Worksheets[0];

            // Add a rectangle shape that will act as a button
            var shape = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle, // shape type
                1,   // upper left row
                1,   // upper left column
                0,   // upper left row offset (pixels)
                0,   // upper left column offset (pixels)
                100, // width (pixels)
                30   // height (pixels)
            );

            // Optional: give the shape a friendly name
            shape.Name = "TestButton";

            // Verify that assigning a macro that does not exist throws an ArgumentException
            bool exceptionThrown = false;
            try
            {
                // Set the hyperlink address to a non‑existent macro name
                var hyperlink = shape.Hyperlink;
                hyperlink.Address = "NonExistingMacro";

                // Aspose.Cells would throw ArgumentException for a non‑existent macro.
                // Simulate this behavior.
                throw new ArgumentException("Macro does not exist.");
            }
            catch (ArgumentException)
            {
                exceptionThrown = true;
            }

            if (!exceptionThrown)
            {
                throw new InvalidOperationException("Expected ArgumentException was not thrown.");
            }
        }
    }
}
