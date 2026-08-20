// Title: Aspose.Cells C# – Detect Worksheet with Data Rows and Shapes (Mixed Content)
// Description: C# example that uses Aspose.Cells to determine whether a worksheet contains at least one non‑blank row (via MaxDisplayRange) and at least one shape (Shapes collection). Returns true only when both data and drawing objects are present.
// Keywords: Aspose.Cells | C# | worksheet data rows | worksheet shapes | mixed content detection | MaxDisplayRange | Shapes.Count | non‑blank row | Excel automation | Aspose.Cells API
// Common Searches: How to check if an Excel sheet has both data and shapes using Aspose.Cells | Aspose.Cells C# detect non‑blank rows and shapes | Get last used row and shape count in Aspose.Cells | Mixed content worksheet validation Aspose.Cells | C# method to verify worksheet contains data rows and drawings
// Developer Intent: Verify that a worksheet includes at least one data row and one shape.
// Use Cases: Skip empty or graphics‑only sheets during batch export | Validate worksheets before publishing to ensure they contain both data and visual elements | Conditionally apply processing only to sheets with mixed content | Generate reports that require both tables and embedded diagrams | Automated quality checks for Excel templates
// AI Prompts: Generate a C# function using Aspose.Cells that returns true when a worksheet has any non‑blank rows and at least one shape. | Rewrite the mixed‑content check to use LINQ for row evaluation while keeping the same logic. | Provide NUnit unit tests for HasDataRowsAndShapes covering empty sheet, data‑only, shape‑only, and both scenarios. | Explain how MaxDisplayRange works in Aspose.Cells and its role in detecting used rows. | Show how to extend the method to also detect charts and pictures.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsMixedContentDemo
{
    // C# example that uses Aspose.Cells to determine whether a worksheet contains at least one non‑blank row (via MaxDisplayRange) and at least one shape (Shapes collection). Returns true only when both data and drawing objects are present.
    public class MixedContentChecker
    {
        // Returns true if the worksheet has at least one non‑blank row and at least one shape.
        public static bool HasDataRowsAndShapes(Worksheet worksheet)
        {
            // Determine the last used row using MaxDisplayRange (returns a Range).
            AsposeRange displayRange = worksheet.Cells.MaxDisplayRange;
            int maxRow = -1;

            if (displayRange != null && displayRange.RowCount > 0)
            {
                // End row = first row + (row count - 1)
                maxRow = displayRange.FirstRow + displayRange.RowCount - 1;
            }

            // If no cells are used, there is nothing to check.
            if (maxRow < 0)
                return false;

            // Scan rows up to the last used row for a non‑blank row.
            bool hasDataRow = false;
            for (int i = 0; i <= maxRow; i++)
            {
                Row row = worksheet.Cells.Rows[i];
                if (row != null && !row.IsBlank)
                {
                    hasDataRow = true;
                    break;
                }
            }

            // Check for at least one shape in the worksheet.
            bool hasShape = worksheet.Shapes.Count > 0;

            // Return true only when both conditions are satisfied.
            return hasDataRow && hasShape;
        }

        // Demonstration of the mixed‑content detection.
        public static void Run()
        {
            try
            {
                // Create a new workbook.
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Add some data to the worksheet (creates a non‑blank row).
                sheet.Cells["A1"].PutValue("Sample Text");
                sheet.Cells["B2"].PutValue(123);

                // Add a shape to the worksheet.
                Shape shape = sheet.Shapes.AddRectangle(2, 2, 0, 0, 100, 50);
                shape.Name = "DemoRectangle";

                // Perform the mixed‑content check.
                bool result = HasDataRowsAndShapes(sheet);
                Console.WriteLine("Worksheet contains both data rows and shapes: " + result);

                // Save the workbook (optional, demonstrates lifecycle usage).
                workbook.Save("MixedContentDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }

    class Program
    {
        static void Main()
        {
            MixedContentChecker.Run();
        }
    }
}
