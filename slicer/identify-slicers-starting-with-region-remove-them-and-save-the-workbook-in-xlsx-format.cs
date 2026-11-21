// Title: Delete slicers that start with "Region" from an Excel workbook using Aspose.Cells for .NET and save as XLSX
// Common Searches: Aspose.Cells C# remove slicer by name prefix | How to delete slicers starting with Region in an Excel file using .NET | Iterate through slicer collection and remove matching slicers with Aspose.Cells | Remove specific slicers and save the workbook as XLSX using Aspose.Cells | C# code to filter slicer collection and delete slicers with a given prefix
// Tags: remove slicers by name Aspose.Cells | slicer collection iteration C# | delete slicers with prefix Region | save workbook as xlsx Aspose.Cells | Aspose.Cells slicer management

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Slicers; // Required for Slicer and SlicerCollection types

// The example loads an existing XLSX workbook, loops through each worksheet's SlicerCollection in reverse order, removes any slicer whose Name begins with "Region" (case‑insensitive), and then saves the updated workbook as an XLSX file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Get the slicer collection for the current worksheet
                SlicerCollection slicers = sheet.Slicers;

                // Iterate backwards to safely remove items while looping
                for (int i = slicers.Count - 1; i >= 0; i--)
                {
                    Slicer slicer = slicers[i];

                    // Remove slicers whose name starts with "Region" (case‑insensitive)
                    if (!string.IsNullOrEmpty(slicer.Name) &&
                        slicer.Name.StartsWith("Region", StringComparison.OrdinalIgnoreCase))
                    {
                        slicers.RemoveAt(i);
                    }
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
