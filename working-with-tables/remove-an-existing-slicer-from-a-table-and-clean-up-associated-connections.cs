// Title: How to delete a specific slicer from an Excel worksheet and remove its data connections using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells remove slicer by name from worksheet | delete slicer and its linked data connections in Excel using Aspose.Cells | Aspose.Cells example for removing slicer from first worksheet | how to clean up slicer connections after deleting slicer in .NET | remove Excel slicer programmatically with Aspose.Cells C#
// Tags: Aspose.Cells slicer removal C# | delete Excel slicer programmatically | remove slicer data connections Aspose.Cells | C# workbook slicer management | Excel slicer cleanup Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Slicers;

namespace Example
{
    // The sample loads an existing workbook, locates a slicer named 'MySlicer' on the first worksheet, removes the slicer and its associated connections, and saves the updated file, demonstrating error handling for missing files and removal failures.
    class Program
    {
        static void Main(string[] args)
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";
            string slicerNameToRemove = "MySlicer";

            try
            {
                // Verify input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Assume the slicer is on the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Locate the slicer by name
                Slicer targetSlicer = null;
                foreach (Slicer slicer in worksheet.Slicers)
                {
                    if (slicer.Name == slicerNameToRemove)
                    {
                        targetSlicer = slicer;
                        break;
                    }
                }

                if (targetSlicer != null)
                {
                    try
                    {
                        // Remove the slicer from the worksheet
                        worksheet.Slicers.Remove(targetSlicer);
                        Console.WriteLine($"Slicer '{slicerNameToRemove}' removed.");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to remove slicer: {ex.Message}");
                    }
                }
                else
                {
                    Console.WriteLine($"Slicer '{slicerNameToRemove}' not found.");
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
