// Title: Delete an embedded OLE object by its label from an Excel worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that finds an OleObject whose Name matches a specific label and removes it from a worksheet. | Show how to iterate the OleObjects collection in reverse order and call RemoveAt to delete a targeted OLE object before saving the workbook. | Demonstrate loading an XLSX file, deleting an unwanted embedded OLE object identified by its label, and persisting the changes with Aspose.Cells.
// Common Searches: asp.net aspose.cells delete ole object by name from worksheet | c# remove specific embedded OLE object from Excel file using Aspose | how to programmatically eliminate unwanted OLE objects in Excel with Aspose.Cells | reverse iteration of OleObjects collection for safe removal aspose.cells example | save workbook after removing OLE shape c# aspose.cells
// Tags: delete oleobject aspose.cells c# | remove oleobject by name worksheet | reverse iteration oleobjects removal | excel oleobject deletion aspnet | save workbook after shape removal

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// The example loads an Excel file, scans the first worksheet's OleObjects collection in reverse, removes any OLE object whose Name matches a given label (case‑insensitive), and saves the updated workbook. It includes file‑existence checking and exception handling.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file '{inputPath}' not found.");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Define the label (name) that identifies the unwanted OLE object
            string unwantedLabel = "UnwantedLabel";

            // Iterate the OLE objects collection in reverse to allow safe removal
            for (int i = sheet.OleObjects.Count - 1; i >= 0; i--)
            {
                OleObject ole = sheet.OleObjects[i];
                if (!string.IsNullOrEmpty(ole.Name) &&
                    ole.Name.Equals(unwantedLabel, StringComparison.OrdinalIgnoreCase))
                {
                    // Remove the matching OLE object
                    sheet.OleObjects.RemoveAt(i);
                }
            }

            // Save the modified workbook
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
