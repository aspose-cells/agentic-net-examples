// Title: Check Shape Adjustment Count Before and After ConvertStringToNumericValue in Aspose.Cells (C#)
// Description: Demonstrates how to record Geometry.ShapeAdjustValues.Count of a Chevron auto‑shape, run sheet.Cells.ConvertStringToNumericValue on a worksheet, and compare the counts to verify that shape adjustment data remains unchanged.
// Keywords: Aspose.Cells | C# | .NET | Shape adjustment values | Geometry.ShapeAdjustValues | ConvertStringToNumericValue | auto shape | Chevron shape | data integrity | worksheet conversion
// Common Searches: Aspose.Cells compare shape adjustment count after converting strings to numbers | C# verify Geometry.ShapeAdjustValues unchanged after ConvertStringToNumericValue | how to ensure auto shape adjustments are preserved in Aspose.Cells | check shape adjustment values count before and after cell conversion | Aspose.Cells data integrity for shape geometry
// Developer Intent: Confirm that converting string cells to numeric values does not modify the number of adjustment values of an auto shape.
// Use Cases: Validate that bulk conversion of worksheet strings to numbers does not affect auto‑shape geometry. | Log a warning when the adjustment values count changes after conversion, indicating a potential integrity issue. | Integrate a safeguard in automated report generation to abort processing if shape adjustments are altered.
// AI Prompts: Write C# code using Aspose.Cells that captures Geometry.ShapeAdjustValues.Count before and after sheet.Cells.ConvertStringToNumericValue and throws an exception if the counts differ. | Show how to log the adjustment values count and handle mismatches when converting worksheet strings to numeric values with Aspose.Cells. | Explain why ConvertStringToNumericValue might impact shape geometry and recommend best practices to protect adjustment values in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // Demonstrates how to record Geometry.ShapeAdjustValues.Count of a Chevron auto‑shape, run sheet.Cells.ConvertStringToNumericValue on a worksheet, and compare the counts to verify that shape adjustment data remains unchanged.
    class CompareShapeAdjustValues
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Add an auto shape (Chevron) which contains adjustment values
                Shape shape = sheet.Shapes.AddAutoShape(AutoShapeType.Chevron, 10, 10, 0, 0, 200, 100);
                Geometry geometry = shape.Geometry;

                // Count adjustment values before any conversion
                int countBefore = geometry.ShapeAdjustValues.Count;
                Console.WriteLine("Adjustment values count before conversion: " + countBefore);

                // Populate cells with string data that can be converted to numeric values
                sheet.Cells["A1"].PutValue("123");
                sheet.Cells["A2"].PutValue("45.6");
                sheet.Cells["A3"].PutValue("NotANumber");

                // Convert all convertible string data in the worksheet to numeric values
                sheet.Cells.ConvertStringToNumericValue();

                // Count adjustment values after conversion
                int countAfter = geometry.ShapeAdjustValues.Count;
                Console.WriteLine("Adjustment values count after conversion: " + countAfter);

                // Compare the counts to ensure data integrity
                if (countBefore == countAfter)
                {
                    Console.WriteLine("Adjustment values count unchanged. Data integrity maintained.");
                }
                else
                {
                    Console.WriteLine("Adjustment values count changed! Potential data integrity issue.");
                }

                // Save the workbook (optional)
                workbook.Save("CompareShapeAdjustValues.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            CompareShapeAdjustValues.Run();
        }
    }
}
