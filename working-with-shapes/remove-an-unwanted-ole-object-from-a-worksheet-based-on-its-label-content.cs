// Title: Remove an OLE object with a given label from an Excel worksheet using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells delete OLE object with specific label | how to remove an embedded OLE object named UnwantedLabel from Excel using Aspose | iterate OleObjectCollection and delete matching objects in Aspose.Cells | programmatically delete OLE objects from a worksheet in .NET | Aspose.Cells example for filtering and removing OLE objects by name
// Tags: remove ole object by label Aspose.Cells | oleobjectcollection reverse iteration C# | delete excel ole object .NET | filter ole objects by name Aspose.Cells | worksheet ole object removal C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample loads an existing workbook, accesses the first worksheet, iterates its OleObjectCollection in reverse, compares each object's Name to a target label (case‑insensitive), removes any matching OLE objects, and saves the modified file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";
        const string unwantedLabel = "UnwantedLabel";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index or name as needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Get the collection of OLE objects on the worksheet
            OleObjectCollection oleObjects = worksheet.OleObjects;

            // Iterate backwards so removal does not affect the loop index
            for (int i = oleObjects.Count - 1; i >= 0; i--)
            {
                OleObject ole = oleObjects[i];

                // Compare the OLE object's name (used as its label) with the target label
                if (!string.IsNullOrEmpty(ole.Name) &&
                    ole.Name.Equals(unwantedLabel, StringComparison.OrdinalIgnoreCase))
                {
                    // Remove the matching OLE object
                    oleObjects.RemoveAt(i);
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
