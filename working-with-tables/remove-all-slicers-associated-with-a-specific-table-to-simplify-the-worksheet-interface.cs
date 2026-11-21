// Title: How to remove all slicers linked to a specific table in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells remove slicers associated with a table name | delete Excel slicers for Table1 programmatically with Aspose.Cells | iterate slicer collection and remove slicers by table in Aspose.Cells .NET | how to clear slicer cache for a specific table using Aspose.Cells | remove table-linked slicers from all worksheets with Aspose.Cells C#
// Tags: Aspose.Cells SlicerCache removal C# | Delete Excel slicers programmatically with Aspose.Cells | Iterate worksheet slicer collection Aspose.Cells | C# remove table-linked slicers from workbook | Aspose.Cells worksheet slicer cleanup

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Slicers; // Required for Slicer and SlicerCollection types

// The code loads an Excel workbook, iterates each worksheet, examines each slicer's SlicerCache name, removes slicers that belong to the specified table, and saves the updated workbook.
class RemoveTableSlicers
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";
        const string targetTableName = "Table1";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Access the slicer collection of the current worksheet
                SlicerCollection slicers = sheet.Slicers;

                // Remove slicers linked to the specified table (iterate backwards)
                for (int i = slicers.Count - 1; i >= 0; i--)
                {
                    Slicer slicer = slicers[i];
                    // Use the slicer's cache name as an identifier (fallback if TableName not available)
                    if (slicer.SlicerCache != null &&
                        string.Equals(slicer.SlicerCache.Name, targetTableName, StringComparison.OrdinalIgnoreCase))
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
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
