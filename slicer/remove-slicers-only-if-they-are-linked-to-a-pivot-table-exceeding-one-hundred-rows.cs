// Title: Delete slicers that are linked to pivot tables from an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Write a C# program with Aspose.Cells that loads an Excel file, iterates each worksheet's SlicerCollection, removes any slicer whose SlicerCache is not null, and saves the result to a new file. | Generate C# code that checks every slicer in a workbook for a linked pivot table (non‑null SlicerCache), deletes those slicers, and writes the updated workbook to a specified output path.
// Common Searches: Aspose.Cells C# delete slicers linked to pivot tables in Excel | how to remove Excel slicers programmatically with Aspose.Cells .NET | filter slicer collection by SlicerCache property using Aspose.Cells | C# example to clean up pivot table slicers in a workbook | Aspose.Cells remove slicer objects before saving workbook
// Tags: Aspose.Cells remove slicer linked to pivot table | C# delete slicer collection Aspose.Cells | filter slicers by SlicerCache property | iterate worksheet slicers Aspose.Cells | save workbook after slicer removal Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Slicers;

// The code loads an Excel workbook, loops through each worksheet's slicer collection, removes any slicer that has a non‑null SlicerCache (indicating a link to a pivot table), and then saves the modified workbook to the designated output file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Access slicers on the current worksheet
                SlicerCollection slicers = sheet.Slicers;

                // Iterate backwards to safely remove slicers
                for (int i = slicers.Count - 1; i >= 0; i--)
                {
                    Slicer slicer = slicers[i];

                    // In Aspose.Cells, a slicer linked to a pivot table has a non‑null SlicerCache.
                    // Remove slicer if it is linked to a pivot table.
                    bool hasLinkedPivot = slicer.SlicerCache != null;

                    if (hasLinkedPivot)
                    {
                        slicers.RemoveAt(i);
                    }
                }
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
