// Title: How to remove a named slicer from every worksheet in an Excel file with Aspose.Cells for .NET
// AI Prompts: Generate C# code that opens an .xlsx file, finds a slicer called 'MySlicer' on any worksheet, deletes it, and saves the workbook as a new file using Aspose.Cells. | Write a method that verifies the input Excel file exists, iterates through all worksheets, removes slicers matching a supplied name, and returns a status message while handling exceptions with Aspose.Cells. | Create a console application that logs which worksheets contained the target slicer, removes those slicers, and writes the updated workbook to a specified output path using Aspose.Cells for .NET.
// Common Searches: aspnet remove slicer by name from Excel workbook using Aspose.Cells | c# delete specific slicer from all sheets in .xlsx with Aspose.Cells | how to programmatically remove slicer to stop pivot refresh in Aspose.Cells | sample code for removing slicer and saving workbook in Aspose.Cells for .NET | Aspose.Cells example to iterate worksheets and delete slicer
// Tags: Aspose.Cells remove slicer by name | C# delete Excel slicer programmatically | iterate worksheets to delete slicer Aspose.Cells | save workbook after slicer removal | handle missing slicer exception Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Slicers;
using System;
using System.IO;

// The example loads an existing workbook, checks that the input .xlsx file is present, iterates through each worksheet, removes any slicer whose Name matches the specified value, logs the removal actions, and saves the modified workbook to a new file while handling potential errors.
class RemoveSlicerPivotConnection
{
    static void Main()
    {
        try
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

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            bool slicerRemoved = false;

            // Iterate through all worksheets and remove slicers with the specified name
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate backwards because we may remove items while iterating
                for (int i = sheet.Slicers.Count - 1; i >= 0; i--)
                {
                    Slicer slicer = sheet.Slicers[i];
                    if (slicer.Name.Equals(slicerName, StringComparison.OrdinalIgnoreCase))
                    {
                        // Remove the slicer from the worksheet
                        sheet.Slicers.RemoveAt(i);
                        slicerRemoved = true;
                        Console.WriteLine($"Removed slicer '{slicerName}' from worksheet '{sheet.Name}'.");
                    }
                }
            }

            if (!slicerRemoved)
            {
                Console.WriteLine($"Slicer '{slicerName}' not found in the workbook.");
            }

            // Save the workbook with the slicer removed
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved as '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
