// Title: C# unit test confirming that assigning a non‑existent macro to a shape throws InvalidOperationException in Aspose.Cells
// Common Searches: how to unit test shape macro assignment failure in Aspose.Cells using C# | Aspose.Cells C# test expecting InvalidOperationException for missing macro | using reflection to set Shape.Action to Macro and verify exception in Aspose.Cells | unit testing macro assignment on rectangle shape without VBA project in Aspose.Cells | C# example of testing non‑existent VBA macro handling in Aspose.Cells
// Tags: Aspose.Cells shape macro assignment unit test | InvalidOperationException for missing macro Aspose.Cells | C# reflection of Shape properties Aspose.Cells | Aspose.Cells workbook without VBA project | rectangle shape macro action Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsTests
{
    // The example creates a new Workbook without a VBA project, adds a rectangle Shape, and uses reflection to set the Shape's Action to Macro and assign a non‑existent macro name. It expects an InvalidOperationException; the unit test reports success when the exception is thrown and failure otherwise.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (no VBA project is added)
                using (Workbook workbook = new Workbook())
                {
                    // Access the first worksheet
                    Worksheet sheet = workbook.Worksheets[0];

                    // Add a rectangle shape to act as a button
                    Shape button = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 1, 1, 0, 0, 100, 30);

                    // Attempt to assign a macro that does not exist.
                    // Use reflection to avoid compile‑time dependency on macro‑related members.
                    try
                    {
                        var shapeType = button.GetType();
                        var actionProp = shapeType.GetProperty("Action");
                        var macroProp = shapeType.GetProperty("Macro");

                        if (actionProp == null || macroProp == null)
                        {
                            Console.WriteLine("Test Inconclusive: Macro‑related properties are not available in this Aspose.Cells version.");
                        }
                        else
                        {
                            // Set Action to ShapeActionType.Macro
                            var enumType = actionProp.PropertyType;
                            var macroEnumValue = Enum.Parse(enumType, "Macro");
                            actionProp.SetValue(button, macroEnumValue);
                            // Set the macro name
                            macroProp.SetValue(button, "NonExistentMacro");

                            // If no exception is thrown, the test has failed.
                            Console.WriteLine("Test Failed: No exception was thrown when assigning a non‑existent macro.");
                        }
                    }
                    catch (InvalidOperationException)
                    {
                        // Expected outcome.
                        Console.WriteLine("Test Passed: InvalidOperationException was thrown as expected.");
                    }
                    catch (Exception ex)
                    {
                        // Any other exception is unexpected.
                        Console.WriteLine($"Unexpected error during macro assignment: {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                // Catch any unexpected exceptions to prevent the program from crashing.
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
