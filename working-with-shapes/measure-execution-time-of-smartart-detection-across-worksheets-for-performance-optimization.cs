// Title: Measure performance of SmartArt detection across all worksheets using Aspose.Cells for .NET
// Common Searches: how to profile SmartArt detection time in an Excel workbook with Aspose.Cells C# | using Stopwatch to benchmark shape iteration in Aspose.Cells worksheets | count SmartArt objects in each worksheet and log execution time with Aspose.Cells | measure total processing time for SmartArt shapes across multiple sheets in .NET | optimize SmartArt detection performance in Aspose.Cells by timing per worksheet
// Tags: SmartArt detection performance Aspose.Cells | worksheet shape iteration timing | Stopwatch profiling Excel shapes | measure SmartArt processing time .NET | shape analysis performance optimization Aspose.Cells

using System;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an Excel workbook with Aspose.Cells, iterates through every worksheet, retrieves each sheet's ShapeCollection, counts shapes where IsSmartArt is true, and uses System.Diagnostics.Stopwatch to record per‑worksheet and total detection times before saving the file and handling I/O errors.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        Workbook workbook = null;
        try
        {
            // Load the workbook
            workbook = new Workbook(inputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to load workbook: {ex.Message}");
            return;
        }

        // Start measuring total execution time
        Stopwatch totalTimer = Stopwatch.StartNew();

        // Iterate through all worksheets in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Optional: measure time per worksheet
            Stopwatch sheetTimer = Stopwatch.StartNew();

            // Get all shapes on the current worksheet
            ShapeCollection shapes = sheet.Shapes;

            int smartArtCount = 0;

            // Detect SmartArt objects among the shapes
            foreach (Shape shape in shapes)
            {
                // Aspose.Cells marks SmartArt shapes with IsSmartArt = true
                if (shape.IsSmartArt)
                {
                    smartArtCount++;
                    // Additional SmartArt processing is not supported directly in Aspose.Cells.
                    // If needed, further handling can be implemented via other libraries.
                }
            }

            sheetTimer.Stop();
            Console.WriteLine($"Worksheet \"{sheet.Name}\" processed in {sheetTimer.ElapsedMilliseconds} ms. SmartArt found: {smartArtCount}");
        }

        totalTimer.Stop();
        Console.WriteLine($"Total SmartArt detection time: {totalTimer.ElapsedMilliseconds} ms.");

        try
        {
            // Ensure the directory for the output file exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook (if any changes were made)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to save workbook: {ex.Message}");
        }
    }
}
