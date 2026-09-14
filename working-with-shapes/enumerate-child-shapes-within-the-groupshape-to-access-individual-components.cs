// Title: How to enumerate child shapes of a GroupShape in an Excel worksheet using Aspose.Cells for .NET
// AI Prompts: Write C# code that scans all shapes on a worksheet, detects GroupShape objects, and attempts to list their contained shapes with Aspose.Cells. | Show a method to retrieve the type and index of each shape inside a GroupShape while handling the Aspose.Cells API limitation. | Provide a C# workaround for accessing individual elements of a grouped shape in an Excel file using Aspose.Cells.
// Common Searches: Aspose.Cells .NET list shapes inside a GroupShape | C# get child shapes of grouped shape in Excel using Aspose.Cells | How to iterate over shapes within a GroupShape with Aspose.Cells for .NET | Aspose.Cells enumerate group shape components programmatically
// Tags: enumerate child shapes Aspose.Cells .NET | grouped shape traversal Excel Aspose.Cells | retrieve group shape components C# | Aspose.Cells shape collection limitation | access individual shapes within GroupShape .NET

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// The example loads an Excel workbook, checks the first shape on the first worksheet, determines if it is a GroupShape, and notes that Aspose.Cells does not directly expose child shapes, prompting the need for workarounds to enumerate individual components before saving the file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook.
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure there is at least one shape.
            if (sheet.Shapes.Count == 0)
            {
                Console.WriteLine("No shapes found on the worksheet.");
                workbook.Save(outputPath);
                return;
            }

            // Retrieve the first shape and check if it is a GroupShape.
            Shape shape = sheet.Shapes[0];
            if (shape is GroupShape group)
            {
                // Aspose.Cells does not expose child shapes directly.
                // If needed, additional processing can be implemented here.
                Console.WriteLine("The first shape is a GroupShape.");
            }
            else
            {
                Console.WriteLine("The first shape is not a GroupShape.");
            }

            // Save the workbook (even if no changes were made).
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
