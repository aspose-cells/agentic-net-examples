// Title: How to delete all slicers linked to a specific table in an Excel worksheet using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# remove slicers that reference a table name | delete slicer objects from a worksheet based on table association in .NET | C# code to clear slicer cache for a given Excel table using Aspose.Cells | remove all slicers tied to a particular table in an Excel file with Aspose.Cells | how to programmatically purge table slicers in a workbook using Aspose.Cells
// Tags: Aspose.Cells remove slicers by table name | C# delete worksheet slicer objects | Aspose.Cells slicer cache removal | Excel worksheet slicer cleanup .NET | programmatic slicer deletion Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Slicers;   // Required for Slicer classes

// The example loads an Excel workbook, selects a worksheet (by name or the first sheet), identifies slicers whose SlicerCache name matches a specified table, removes those slicers from the worksheet, and saves the updated workbook to the output path.
class Program
{
    static void Main()
    {
        // Define input and output paths (replace placeholders with actual paths)
        string inputPath = "{InputFilePath}";
        string outputPath = "{OutputFilePath}";
        string worksheetName = "{WorksheetName}";
        string targetTableName = "{TableName}";

        try
        {
            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input file not found: {inputPath}");

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the worksheet by name (or fallback to the first worksheet)
            Worksheet sheet = null;
            if (!string.IsNullOrEmpty(worksheetName) && workbook.Worksheets[worksheetName] != null)
                sheet = workbook.Worksheets[worksheetName];
            else
                sheet = workbook.Worksheets[0];

            if (sheet == null)
                throw new InvalidOperationException("Worksheet could not be determined.");

            // Collect indices of slicers linked to the specified table
            List<int> slicerIndices = new List<int>();
            for (int i = 0; i < sheet.Slicers.Count; i++)
            {
                Slicer slicer = sheet.Slicers[i];
                // Use the SlicerCache name (or other identifiable property) to match the table
                string slicerCacheName = slicer.SlicerCache?.Name;
                if (!string.IsNullOrEmpty(slicerCacheName) &&
                    string.Equals(slicerCacheName, targetTableName, StringComparison.OrdinalIgnoreCase))
                {
                    slicerIndices.Add(i);
                }
            }

            // Remove slicers in reverse order to avoid index shifting
            for (int i = slicerIndices.Count - 1; i >= 0; i--)
                sheet.Slicers.RemoveAt(slicerIndices[i]);

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine("Workbook saved successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
