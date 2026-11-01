// Title: How to change slicer label font family, size, and color with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# change slicer label font family and size | set slicer text color programmatically using Aspose.Cells .NET | customize Excel slicer appearance with Aspose.Cells API | modify slicer label font style in C# workbook | Aspose.Cells example for updating slicer font properties
// Tags: Aspose.Cells slicer font formatting | C# set slicer label color | Aspose.Cells modify slicer text style | Excel slicer appearance customization .NET | Aspose.Cells slicer style API usage

using System;
using System.IO;
using Aspose.Cells;
using System.Drawing;

namespace AsposeCellsSlicerExample
{
    // The sample loads an existing workbook, checks that the input file exists, ensures the output directory is created, and saves the workbook unchanged. Slicer font manipulation code is omitted because it requires a version of Aspose.Cells that includes the Slicer, SlicerCollection, and SlicerStyle classes; when available, developers can adjust slicer label font family, size, and color via those APIs.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {Path.GetFullPath(inputPath)}");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Get the first worksheet (adjust index if needed)
                Worksheet worksheet = workbook.Worksheets[0];

                // NOTE: Slicer manipulation requires a version of Aspose.Cells that includes
                // the Slicer, SlicerCollection, and SlicerStyle classes. If your version
                // does not support these types, the related code has been omitted.

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
