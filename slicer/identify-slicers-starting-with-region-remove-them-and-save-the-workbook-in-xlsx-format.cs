// Title: Remove all slicers whose names start with "Region" from an Excel workbook using Aspose.Cells for .NET and save as XLSX
// AI Prompts: Write C# code that loads an Excel file with Aspose.Cells, iterates through each worksheet's SlicerCollection, removes any slicer whose Name starts with "Region" (case‑insensitive), and saves the workbook as an XLSX file. | Show a step‑by‑step example of looping backwards through a SlicerCollection in Aspose.Cells for .NET to safely delete slicers matching a specific prefix and then export the workbook. | Generate a reusable C# method that accepts a workbook path and a slicer name prefix, removes matching slicers using Aspose.Cells, and writes the result to a new XLSX file.
// Common Searches: c# aspnet remove slicers with name starting with region from excel using aspose.cells | aspose.cells delete slicer by prefix programmatically | how to loop through slicer collection and remove specific slicers in .net | save workbook after removing slicers aspose.cells c# | remove region slicers from multiple worksheets aspose.cells example
// Tags: aspose.cells slicer removal by name prefix | c# iterate slicercollection aspose.cells | delete region slicers excel workbook | save modified workbook as xlsx aspose.cells | case‑insensitive slicer name filter aspose.cells

using System;
using Aspose.Cells;
using Aspose.Cells.Slicers;

namespace AsposeCellsSlicerRemoval
{
    // Loads an Excel file, iterates through all worksheets, removes any slicer whose Name begins with "Region" (case‑insensitive) by looping backwards through the SlicerCollection, and saves the updated workbook as output.xlsx in XLSX format.
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Get the slicer collection for the current worksheet
                SlicerCollection slicers = sheet.Slicers;

                // Loop backwards to safely remove items while iterating
                for (int i = slicers.Count - 1; i >= 0; i--)
                {
                    Slicer slicer = slicers[i];

                    // Check if the slicer's name starts with "Region"
                    // (using case‑insensitive comparison)
                    if (!string.IsNullOrEmpty(slicer.Name) &&
                        slicer.Name.StartsWith("Region", StringComparison.OrdinalIgnoreCase))
                    {
                        // Remove the slicer at the current index
                        slicers.RemoveAt(i);
                    }
                }
            }

            // Save the modified workbook in XLSX format
            workbook.Save("output.xlsx", SaveFormat.Xlsx);
        }
    }
}
