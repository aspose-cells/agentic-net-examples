// Title: Verify ShapeAdjustValues Count Before and After Workbook Save/Load with Aspose.Cells for .NET
// Description: Creates a workbook, adds a Chevron auto shape, records the number of ShapeAdjustValues, optionally changes a value, saves the file, reloads it, and compares the counts to ensure shape geometry integrity.
// Keywords: Aspose.Cells ShapeAdjustValues | auto shape adjustment count | C# workbook save reload verification | shape geometry persistence | data integrity Aspose.Cells
// Common Searches: Aspose.Cells check shape adjustment values after save | compare ShapeAdjustValues count before and after reload C# | verify auto shape geometry persistence in Excel | shape adjust values lost after conversion Aspose.Cells | how to test shape data integrity with Aspose.Cells
// Developer Intent: Confirm that the count of ShapeAdjustValues remains unchanged after a workbook is saved and reloaded.
// Use Cases: Automated regression test for shape geometry retention in generated Excel files. | Detecting loss of custom auto‑shape parameters during format conversion (e.g., XLSX → PDF). | Ensuring shape data consistency in document‑generation pipelines that involve multiple save/load cycles.
// AI Prompts: Write C# code that asserts ShapeAdjustValues.Count is identical before saving and after loading a workbook with Aspose.Cells. | Provide a method to log differences in ShapeAdjustValues when a workbook is reloaded. | Explain how to modify a shape's adjustment values and verify they stay unchanged after converting the workbook to another format.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // Creates a workbook, adds a Chevron auto shape, records the number of ShapeAdjustValues, optionally changes a value, saves the file, reloads it, and compares the counts to ensure shape geometry integrity.
    class CompareShapeAdjustValues
    {
        static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add an auto shape (Chevron) to the worksheet
            Shape shape = worksheet.Shapes.AddAutoShape(AutoShapeType.Chevron, 10, 10, 0, 0, 200, 100);
            Geometry geometry = shape.Geometry;

            // Record the number of adjustment values before any conversion
            int initialAdjustCount = geometry.ShapeAdjustValues.Count;

            // Optionally modify the first adjust value if any exist
            if (initialAdjustCount > 0)
            {
                geometry.ShapeAdjustValues[0].Value = 0.3;
            }

            // Save the workbook (lifecycle step)
            string filePath = "AdjustValuesBeforeAfter.xlsx";
            workbook.Save(filePath);

            // Ensure the file exists before loading
            if (!File.Exists(filePath))
            {
                Console.WriteLine("Saved file not found.");
                return;
            }

            // Load the workbook again to simulate a conversion/reload scenario
            Workbook loadedWorkbook = new Workbook(filePath);
            Worksheet loadedWorksheet = loadedWorkbook.Worksheets[0];
            Shape loadedShape = loadedWorksheet.Shapes[0];
            Geometry loadedGeometry = loadedShape.Geometry;

            // Record the number of adjustment values after reload
            int afterAdjustCount = loadedGeometry.ShapeAdjustValues.Count;

            // Compare the counts and output the result
            Console.WriteLine($"Initial adjustment values count: {initialAdjustCount}");
            Console.WriteLine($"After reload adjustment values count: {afterAdjustCount}");
            Console.WriteLine(initialAdjustCount == afterAdjustCount
                ? "Adjustment values count unchanged – data integrity preserved."
                : "Adjustment values count changed – data integrity issue.");
        }
    }
}
