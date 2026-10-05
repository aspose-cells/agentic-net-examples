// Title: Remove a specific slicer named "MySlicer" from all worksheets in an XLSX file using Aspose.Cells for .NET and save the result
// AI Prompts: Write C# code that opens an existing XLSX workbook with Aspose.Cells, locates the slicer whose Name is "MySlicer" on every worksheet, deletes it, and saves the workbook as a new file. | Generate a .NET snippet that checks the input file, iterates each sheet's SlicerCollection in reverse order, removes the matching slicer, handles possible exceptions, and writes the updated workbook.
// Common Searches: Aspose.Cells C# remove slicer by name from Excel workbook | How to delete a specific slicer from all sheets using Aspose.Cells .NET | Programmatically remove named slicer in XLSX with Aspose.Cells | C# iterate worksheets and delete slicer MySlicer | Save workbook after removing slicer using Aspose.Cells
// Tags: aspocells remove slicer from xlsx | c# slicercollection delete by name | aspocells workbook slicer manipulation | excel slicer removal using .net | save workbook after slicer deletion

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Slicers;

// The example verifies the input XLSX file, loads it with Aspose.Cells, loops through each worksheet's SlicerCollection in reverse, removes any slicer whose Name matches "MySlicer", and saves the modified workbook as a new XLSX file while handling exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";
        const string slicerName = "MySlicer";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file '{inputPath}' not found.");
            return;
        }

        try
        {
            // Load the existing XLSX workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets to find and remove the slicer
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                SlicerCollection slicers = sheet.Slicers;

                // Iterate backwards when removing items from a collection
                for (int i = slicers.Count - 1; i >= 0; i--)
                {
                    if (slicers[i].Name == slicerName)
                    {
                        // Remove the slicer from the collection
                        slicers.RemoveAt(i);
                    }
                }
            }

            // Save the modified workbook as XLSX
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
