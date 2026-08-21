// Title: C# – Remove “Region*” slicers from an Excel workbook with Aspose.Cells and save as XLSX
// Description: Loads a workbook, scans each worksheet’s SlicerCollection, deletes slicers whose Name starts with “Region” (case‑insensitive) using reverse iteration, and saves the modified file as an XLSX document.
// Keywords: Aspose.Cells | C# | remove slicer | slicer collection | delete slicer by name | Region slicer | save workbook xlsx | Excel slicer removal | Aspose.Cells API | batch workbook processing
// Common Searches: Aspose.Cells delete slicer by prefix | C# remove Excel slicer starting with Region | iterate SlicerCollection Aspose.Cells | save workbook after slicer removal Aspose | remove all slicers from workbook using Aspose.Cells
// Developer Intent: Programmatically eliminate every slicer whose name begins with “Region” from all worksheets and output the cleaned workbook in XLSX format.
// Use Cases: Strip automatically generated region slicers before publishing a report to stakeholders. | Batch‑process a set of workbooks to clean up temporary slicers after a data refresh. | Prepare an archive copy of a workbook by removing slicers that are no longer relevant.
// AI Prompts: Generate C# code with Aspose.Cells that deletes slicers whose names start with a given prefix and then saves the workbook as XLSX. | Explain how to safely iterate a SlicerCollection in reverse order to avoid index errors when removing matching slicers. | Create a reusable method that accepts a file path and a slicer name prefix, removes matching slicers from every worksheet, and writes the result to a new file.

using System;
using Aspose.Cells;
using Aspose.Cells.Slicers;

namespace AsposeCellsSlicerRemoval
{
    // Loads a workbook, scans each worksheet’s SlicerCollection, deletes slicers whose Name starts with “Region” (case‑insensitive) using reverse iteration, and saves the modified file as an XLSX document.
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

                // Remove slicers whose name starts with "Region"
                // Iterate backwards to avoid index shifting when removing
                for (int i = slicers.Count - 1; i >= 0; i--)
                {
                    Slicer slicer = slicers[i];
                    if (slicer.Name != null && slicer.Name.StartsWith("Region", StringComparison.OrdinalIgnoreCase))
                    {
                        slicers.RemoveAt(i);
                    }
                }
            }

            // Save the modified workbook in XLSX format
            workbook.Save("output.xlsx", SaveFormat.Xlsx);
        }
    }
}
