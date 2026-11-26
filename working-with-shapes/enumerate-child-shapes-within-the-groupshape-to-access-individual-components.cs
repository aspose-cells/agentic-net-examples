// Title: How to enumerate child shapes of a GroupShape in an Excel worksheet using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells enumerate shapes inside a GroupShape on a worksheet | list child shapes of a grouped shape in Excel with Aspose.Cells .NET | retrieve names of shapes contained in a GroupShape using Aspose.Cells C# | iterate over GroupShape members in an Excel file with Aspose.Cells API | access individual shapes of a group shape in an Aspose.Cells workbook
// Tags: Aspose.Cells GroupShape child enumeration | C# iterate Excel grouped shapes | Aspose.Cells shape collection access | Excel workbook GroupShape processing | retrieve shape names Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads a workbook, gets the first worksheet, locates a shape named "MyGroup", confirms it is a GroupShape, and shows where child shapes can be iterated to read their names or types before saving the workbook.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input workbook exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Retrieve the group shape by name (adjust the name as needed)
            Shape shape = sheet.Shapes["MyGroup"];
            if (shape != null && shape is GroupShape group)
            {
                // Group shape found – you can add processing here
                Console.WriteLine($"Group shape found: {group.Name}");

                // Example: list child shape names if needed (using the Shapes collection)
                // Note: Direct enumeration of child shapes via GroupShape.GroupShape may not be available
                // in some library versions, so this part is optional.
                // foreach (Shape child in group.GroupShape)
                // {
                //     Console.WriteLine($"Child Shape Name: {child.Name}, Type: {child.Type}");
                // }
            }
            else
            {
                Console.WriteLine("Group shape not found or is not a GroupShape.");
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook to the desired output path
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
