// Title: How to cast a worksheet Shape to GroupShape and verify its type using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loops through all shapes on the first worksheet, uses the 'as' operator to cast each to GroupShape, and writes a confirmation message when the cast succeeds. | Show a safe way to identify a GroupShape among worksheet shapes and access its properties with Aspose.Cells in a .NET project. | Provide a C# example that loads an Excel file, checks for the presence of a GroupShape, casts it, performs a simple manipulation, and then saves the workbook.
// Common Searches: Aspose.Cells C# find and cast GroupShape from worksheet shapes | how to check if a shape is a GroupShape in Aspose.Cells .NET | C# code to iterate Excel shapes and identify group shapes using Aspose.Cells
// Tags: group shape casting Aspose.Cells C# | identify groupshape in worksheet Aspose.Cells | shape type verification Aspose.Cells .NET | iterate worksheet shapes Aspose.Cells | group shape manipulation example Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsGroupShapeDemo
{
    // The sample loads an Excel workbook, ensures the first worksheet contains shapes, iterates through them to locate a GroupShape, casts the shape using the 'as' operator, confirms the cast succeeded, optionally manipulates the group, and saves the workbook.
    class Program
    {
        static void Main()
        {
            try
            {
                const string inputPath = "input.xlsx";
                const string outputPath = "output.xlsx";

                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {Path.GetFullPath(inputPath)}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet (ensure at least one worksheet exists)
                if (workbook.Worksheets.Count == 0)
                {
                    Console.WriteLine("The workbook does not contain any worksheets.");
                    return;
                }

                Worksheet worksheet = workbook.Worksheets[0];

                // Ensure the worksheet contains at least one shape
                if (worksheet.Shapes.Count == 0)
                {
                    Console.WriteLine("No shapes were found on the first worksheet.");
                    return;
                }

                // Attempt to find a GroupShape among the shapes
                GroupShape groupShape = null;
                foreach (Shape shape in worksheet.Shapes)
                {
                    groupShape = shape as GroupShape;
                    if (groupShape != null)
                    {
                        break; // Found a GroupShape, exit loop
                    }
                }

                if (groupShape != null)
                {
                    // The shape is a GroupShape – you can now work with groupShape
                    Console.WriteLine("The shape has been successfully cast to GroupShape.");
                    // Example: you could manipulate the groupShape here
                }
                else
                {
                    // No GroupShape found – handle as needed
                    Console.WriteLine("No GroupShape was found among the shapes on the worksheet.");
                }

                // Save the workbook if any modifications were made (optional)
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                // Catch any unexpected exceptions and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
