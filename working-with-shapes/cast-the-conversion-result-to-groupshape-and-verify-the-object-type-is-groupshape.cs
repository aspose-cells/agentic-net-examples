// Title: How to cast a worksheet shape to GroupShape and verify its type using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells cast shape to GroupShape and test if it is a group | Determine whether a worksheet shape is a GroupShape in Aspose.Cells .NET | Aspose.Cells shape type checking example C# | Convert worksheet shape to GroupShape with Aspose.Cells and verify cast | Using Aspose.Cells drawing API to identify group shapes in an Excel file
// Tags: Aspose.Cells GroupShape casting C# | worksheet shape type verification Aspose.Cells | Aspose.Cells drawing API group shape detection | C# Excel shape to GroupShape conversion | Aspose.Cells shape as operator usage

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // The example loads an Excel workbook, retrieves the first shape on the first worksheet, attempts to cast it to a GroupShape using the C# 'as' operator, and prints whether the cast succeeded, demonstrating shape type verification with Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            string filePath = "sample.xlsx";

            try
            {
                // Verify the input file exists
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found: {filePath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(filePath);

                // Ensure there is at least one worksheet
                if (workbook.Worksheets.Count == 0)
                {
                    Console.WriteLine("The workbook contains no worksheets.");
                    return;
                }

                Worksheet sheet = workbook.Worksheets[0];

                // Ensure there is at least one shape on the worksheet
                if (sheet.Shapes.Count == 0)
                {
                    Console.WriteLine("No shapes found in the first worksheet.");
                    return;
                }

                // Retrieve the first shape
                Shape shape = sheet.Shapes[0];

                // Attempt to cast the shape to GroupShape
                GroupShape groupShape = shape as GroupShape;

                if (groupShape != null)
                {
                    Console.WriteLine("The shape is a GroupShape.");
                }
                else
                {
                    Console.WriteLine("The shape is NOT a GroupShape.");
                }
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
