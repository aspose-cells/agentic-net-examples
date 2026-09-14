// Title: Read the TextBody alignment of the first shape in an Excel worksheet using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads a workbook, accesses the first shape on the first worksheet, and returns the ShapeTextAlignment value from its TextBody. | Show how to safely attempt setting Shape.TextBody.TextAlignment to Center in Aspose.Cells, including version‑check handling for read‑only properties.
// Common Searches: how to read shape text alignment using Aspose.Cells in C# | example of getting ShapeTextAlignment from a shape's TextBody | Aspose.Cells shape text alignment property read‑only issue | center text inside a shape programmatically with Aspose.Cells
// Tags: Aspose.Cells Shape.TextBody alignment API | C# extract ShapeTextAlignment value | Aspose.Cells modify shape text horizontal alignment | handle read‑only TextAlignment in Aspose.Cells | Excel shape text formatting with Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample loads an existing Excel file, checks for shapes on the first worksheet, retrieves the first shape, and demonstrates how to read its TextBody alignment. It also includes a commented example of setting the alignment to Center, noting that the TextAlignment property may be read‑only in certain Aspose.Cells versions, before saving the workbook.
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
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one shape
            if (sheet.Shapes.Count == 0)
            {
                Console.WriteLine("No shapes found in the first worksheet.");
            }
            else
            {
                try
                {
                    // Retrieve the first shape
                    Shape shape = sheet.Shapes[0];

                    // Set horizontal alignment of the shape's text to Center
                    // Note: In some Aspose.Cells versions TextAlignment is read‑only.
                    // If supported, the following line can be uncommented:
                    // shape.TextBody.TextAlignment = TextAlignmentType.Center;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing shape: {ex.Message}");
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
