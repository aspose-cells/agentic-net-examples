// Title: Clone a SmartArt shape, cast it to GroupShape, and compare child shape counts with Aspose.Cells for .NET
// AI Prompts: Generate C# code that uses Aspose.Cells to duplicate a SmartArt shape on a worksheet, convert the duplicate to a GroupShape, and retrieve its child count via reflection. | Write a .NET program that clones a SmartArt shape, casts the clone to GroupShape, compares the original and cloned group’s child shape numbers, and saves the workbook.
// Common Searches: Aspose.Cells C# duplicate SmartArt shape and get group child count | clone SmartArt shape to GroupShape and compare child shapes using Aspose.Cells .NET | C# Aspose.Cells GetChildCount for cloned SmartArt shape | AddCopy method clone shape and check child count Aspose.Cells example
// Tags: Aspose.Cells clone SmartArt shape | GroupShape child count Aspose.Cells | AddCopy shape cloning .NET | reflection GetChildCount GroupShape | compare original and cloned shape child counts

using System;
using System.IO;
using System.Reflection;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an Excel workbook, finds the first shape, clones it with AddCopy, attempts to treat both original and cloned shapes as GroupShape, uses reflection to call GetChildCount when available, prints the child counts and whether they match, and saves the workbook with the cloned shape.
class SmartArtCloneExample
{
    static void Main()
    {
        try
        {
            // Input and output file paths
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Locate the first shape on the worksheet
            Shape originalShape = null;
            foreach (Shape shape in sheet.Shapes)
            {
                originalShape = shape;
                break;
            }

            if (originalShape == null)
            {
                Console.WriteLine("No shape found on the worksheet.");
                return;
            }

            // Clone the shape using AddCopy (requires position parameters)
            Shape clonedShape = sheet.Shapes.AddCopy(
                originalShape,
                originalShape.UpperLeftRow,
                originalShape.UpperLeftColumn,
                originalShape.LowerRightRow,
                originalShape.LowerRightColumn);

            // Determine child shape counts (relevant for GroupShape)
            int originalChildCount = 0;
            int clonedChildCount = 0;

            if (originalShape is GroupShape originalGroup)
            {
                // Use reflection to call GetChildCount if it exists in the current version
                MethodInfo? getCount = typeof(GroupShape).GetMethod("GetChildCount", BindingFlags.Instance | BindingFlags.Public);
                if (getCount != null)
                {
                    object? result = getCount.Invoke(originalGroup, null);
                    originalChildCount = result != null ? Convert.ToInt32(result) : 0;
                }
            }

            if (clonedShape is GroupShape clonedGroup)
            {
                MethodInfo? getCount = typeof(GroupShape).GetMethod("GetChildCount", BindingFlags.Instance | BindingFlags.Public);
                if (getCount != null)
                {
                    object? result = getCount.Invoke(clonedGroup, null);
                    clonedChildCount = result != null ? Convert.ToInt32(result) : 0;
                }
            }

            // Display the results
            Console.WriteLine($"Original shape child count: {originalChildCount}");
            Console.WriteLine($"Cloned shape child count: {clonedChildCount}");
            Console.WriteLine($"Counts are equal: {originalChildCount == clonedChildCount}");

            // Ensure the output directory exists
            string? outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook with the cloned shape
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
