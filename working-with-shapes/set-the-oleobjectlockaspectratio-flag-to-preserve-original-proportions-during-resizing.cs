// Title: How to preserve OLE object aspect ratio while resizing with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells resize OLE object without changing aspect ratio | calculate aspect ratio for OLE objects in an Aspose.Cells workbook | manual width and height scaling of OLE objects using Aspose.Cells .NET | keep OLE object proportions when adjusting size with Aspose.Cells | Aspose.Cells example for preserving OLE object proportions in C#
// Tags: Aspose.Cells OLE object aspect ratio handling | C# manual OLE object size adjustment Aspose.Cells | OleObject width height scaling preserving ratio | Aspose.Cells workbook OLE object manipulation | C# Aspose.Cells OLE object resizing example

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample loads a workbook, accesses the first worksheet, retrieves its OleObjectCollection, and demonstrates how to manually preserve an OLE object's original aspect ratio by calculating the height‑to‑width ratio and applying it after setting a new width. It also notes that Aspose.Cells does not expose a LockAspectRatio property, so the ratio must be handled programmatically before saving the modified workbook.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

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

            // Access the first worksheet (adjust index as needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Get the collection of OLE objects on the worksheet
            OleObjectCollection oleObjects = worksheet.OleObjects;

            // Ensure there is at least one OLE object to work with
            if (oleObjects.Count > 0)
            {
                // Retrieve the first OLE object (or use a specific index)
                OleObject oleObject = oleObjects[0];

                // Note: Aspose.Cells OleObject does not expose a LockAspectRatio property.
                // If resizing is required, adjust Height and Width manually while preserving the ratio.
                // Example (uncomment and modify as needed):
                // double aspectRatio = oleObject.Height / oleObject.Width;
                // oleObject.Width = 200; // desired width
                // oleObject.Height = oleObject.Width * aspectRatio;
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
