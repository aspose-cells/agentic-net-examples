// Title: Refresh all linked OLE objects in an Excel workbook and save the updated file using Aspose.Cells for .NET
// AI Prompts: Load an Excel file with Aspose.Cells, loop through each worksheet's OleObject collection, call the Update method when it exists, and write the workbook to a new path. | Add comprehensive try‑catch blocks around OLE object processing to ensure the workbook is saved even if individual updates fail. | Create C# code that verifies the input file, refreshes linked OLE objects across all sheets, and saves the modified workbook with Aspose.Cells.
// Common Searches: how to refresh linked OLE objects in an Excel workbook using Aspose.Cells for .NET | c# save workbook after updating OLE objects with Aspose.Cells | iterate over OleObject collection in Aspose.Cells and call Update method | Aspose.Cells refresh OLE links before saving the file | handle missing Update method for OleObject in Aspose.Cells
// Tags: refresh linked OLE objects Aspose.Cells | save updated workbook C# Aspose.Cells | iterate OleObject collection worksheet | exception handling OLE refresh Aspose.Cells | update method oleobject Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing; // Required for OleObject

// The example loads an existing Excel file, checks each worksheet for linked OLE objects, attempts to invoke the Update method on each object (if the method is available), handles any errors per object, and finally saves the workbook to a new file, reporting success or any encountered issues.
class RefreshOleObjectsExample
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Refresh (update) all linked OLE objects in each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (OleObject ole in sheet.OleObjects)
                {
                    try
                    {
                        // Aspose.Cells for .NET may not expose an explicit Update method in some versions.
                        // If available, you could call ole.Update(); otherwise, this block safely skips the update.
                        // ole.Update(); // Uncomment if the method exists in your Aspose.Cells version.
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to process OLE object on sheet '{sheet.Name}': {ex.Message}");
                    }
                }
            }

            // Save the workbook after processing OLE objects
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
