// Title: How to remove a specific slicer from an XLSX workbook using Aspose.Cells for .NET and save the updated file
// Common Searches: Aspose.Cells C# remove slicer named MySlicer from Excel workbook | Delete a slicer by name in a .xlsx file with Aspose.Cells .NET | C# code to iterate worksheets and remove a slicer using Aspose.Cells | Save workbook after removing slicer with Aspose.Cells SaveFormat.Xlsx | Handle missing input.xlsx when deleting a slicer in Aspose.Cells
// Tags: Aspose.Cells remove slicer from workbook | C# delete named slicer in Excel file | save modified workbook as xlsx Aspose.Cells | iterate worksheet slicer collection C# | handle missing input file Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Slicers; // Required for Slicer class

// The example loads an existing XLSX file, searches each worksheet's Slicers collection for a slicer with the specified name, removes the matching slicer, and then saves the workbook back to XLSX format, including error handling for missing files and other exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";
        const string slicerName = "MySlicer";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"The input file '{inputPath}' was not found.");

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through worksheets to locate and remove the specified slicer
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Search for the slicer by name within the current worksheet
                foreach (Slicer slicer in sheet.Slicers)
                {
                    if (string.Equals(slicer.Name, slicerName, StringComparison.OrdinalIgnoreCase))
                    {
                        // Remove the slicer from the worksheet
                        sheet.Slicers.Remove(slicer);
                        // Assuming slicer names are unique, exit loops after removal
                        goto SaveAndExit;
                    }
                }
            }

        SaveAndExit:
            // Save the modified workbook
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
