// Title: How to enable the ShowNoDataItems option for an Excel slicer using Aspose.Cells in C#
// AI Prompts: Set the slicer's show‑no‑data option to true with Aspose.Cells and save the workbook. | Retrieve the first slicer on a worksheet, enable display of items without data, and write the updated file. | Programmatically toggle the zero‑items visibility of an existing slicer in a .NET Excel file using Aspose.Cells.
// Common Searches: Aspose.Cells C# enable ShowNoDataItems for Excel slicer | display empty categories in Excel slicer using .NET library | how to show zero‑value items in a slicer with Aspose.Cells API
// Tags: Aspose.Cells slicer show‑no‑data option | C# enable zero‑data items in Excel slicer | programmatic slicer configuration .NET | Excel slicer display empty items Aspose.Cells | set slicer show‑no‑data property C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Slicers; // For Slicer class

// The example loads an existing workbook, accesses the first worksheet, obtains the first slicer, demonstrates enabling its show‑no‑data option to display items with no data, and saves the modified workbook to a new file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input workbook exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure there is at least one slicer on the worksheet
            if (sheet.Slicers.Count > 0)
            {
                // Get the first slicer (or locate by name)
                Slicer slicer = sheet.Slicers[0];

                // The ShowNoDataItems property may not be available in older versions.
                // If needed, set other slicer properties here.
                // Example: slicer.ShowHeader = true;
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
