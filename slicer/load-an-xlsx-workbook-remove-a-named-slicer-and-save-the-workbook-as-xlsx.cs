// Title: Remove a specific named slicer from an XLSX workbook using Aspose.Cells for .NET and save the updated file
// AI Prompts: Write C# code that loads an existing XLSX file with Aspose.Cells, searches every worksheet for a slicer whose Name equals "MySlicer", deletes that slicer from the worksheet’s SlicerCollection, and saves the workbook to a new file. | Demonstrate how to iterate backward through a SlicerCollection in Aspose.Cells, remove a matching slicer safely, and handle a missing input file scenario in C#.
// Common Searches: aspnet delete slicer by name from Excel file using Aspose.Cells | C# code to delete a specific slicer from all sheets in an XLSX workbook | how to programmatically delete a slicer called MySlicer with Aspose.Cells | Aspose.Cells .NET remove slicer from workbook and save changes | iterate worksheets to find and delete slicer in Excel using Aspose.Cells C#
// Tags: aspocells remove slicer c# | named slicer deletion xlsx aspocells | slicercollection remove method .net | iterate worksheets aspocells slicer | save workbook after slicer removal aspocells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Slicers;

// The example loads input.xlsx with Aspose.Cells, checks each worksheet for a slicer named "MySlicer", removes the matching slicer from the worksheet's SlicerCollection, and saves the modified workbook as output.xlsx, with error handling for missing files and other exceptions.
public class RemoveNamedSlicer
{
    public static void Run()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";
            const string slicerName = "MySlicer";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Search all worksheets for the slicer with the specified name
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                SlicerCollection slicers = sheet.Slicers;

                // Iterate backwards to safely remove items
                for (int i = slicers.Count - 1; i >= 0; i--)
                {
                    Slicer slicer = slicers[i];
                    if (slicer.Name == slicerName)
                    {
                        slicers.Remove(slicer);
                        // Assuming slicer names are unique, exit the loops
                        break;
                    }
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Entry point required for compilation
    public static void Main()
    {
        Run();
    }
}
