// Title: Identify Excel worksheets with more than ten shapes using Aspose.Cells for .NET
// Common Searches: C# code to find worksheets that have over 10 shapes with Aspose.Cells | how to count shapes in each Excel sheet using Aspose.Cells .NET | list worksheets containing complex graphics in an Excel workbook Aspose.Cells | Aspose.Cells shape collection threshold example C# | detect Excel sheets with high shape count using Aspose.Cells library
// Tags: enumerate worksheet shapes Aspose.Cells | filter sheets by number of shapes | Aspose.Cells shape limit detection | identify worksheets with many graphics | C# iterate Excel shapes per sheet

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace ShapeCountChecker
{
    // The example loads a workbook, iterates through all worksheets, retrieves each sheet's ShapeCollection, counts the shapes, and prints the worksheet name and shape count for any sheet that contains more than ten shapes, indicating complex graphical content.
    class Program
    {
        static void Main(string[] args)
        {
            // Validate input arguments
            if (args.Length == 0)
            {
                Console.WriteLine("Usage: ShapeCountChecker <workbook_path>");
                return;
            }

            string workbookPath = args[0];

            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"Error: File not found - {workbookPath}");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(workbookPath);

                bool anyComplexSheet = false;

                // Iterate through all worksheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Get the ShapeCollection for the current worksheet
                    ShapeCollection shapes = sheet.Shapes;

                    // Count the shapes
                    int shapeCount = shapes.Count;

                    // Check if the count exceeds the threshold
                    if (shapeCount > 10)
                    {
                        anyComplexSheet = true;
                        Console.WriteLine($"Worksheet \"{sheet.Name}\" contains {shapeCount} shapes (complex graphical content).");
                    }
                }

                if (!anyComplexSheet)
                {
                    Console.WriteLine("No worksheets contain more than 10 shapes.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred while processing the workbook: {ex.Message}");
            }
        }
    }
}
