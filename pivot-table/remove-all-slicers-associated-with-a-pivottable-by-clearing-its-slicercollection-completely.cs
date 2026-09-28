// Title: Delete all slicers associated with PivotTables in every worksheet using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an Excel file with Aspose.Cells, iterates through each worksheet, and invokes the slicer collection clear operation. | Show a method to purge slicer objects from a workbook without altering other worksheet data using the Aspose.Cells API. | Provide a snippet that saves the workbook after removing slicers from all sheets in a .NET application.
// Common Searches: Aspose.Cells C# clear slicer collection for all worksheets | remove pivot table slicers programmatically in .NET | how to delete slicers from Excel workbook using Aspose.Cells library
// Tags: Aspose.Cells slicer removal API | remove slicer objects from workbook C# | delete worksheet slicers via Aspose.Cells | clear all slicers in Excel file .NET | programmatic slicer cleanup Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Slicers;
using Aspose.Cells.Pivot;

namespace RemoveAllSlicersDemo
{
    // The example loads an existing Excel workbook, loops through each worksheet, obtains its SlicerCollection, clears all slicers when present, and saves the modified file to a new location.
    class Program
    {
        static void Main(string[] args)
        {
            // Load an existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Get the slicer collection for the current worksheet
                SlicerCollection slicers = sheet.Slicers;

                // If there are any slicers, clear them all
                if (slicers.Count > 0)
                {
                    slicers.Clear(); // Removes all slicers from the worksheet
                }
            }

            // Save the modified workbook (replace with your desired output path)
            workbook.Save("output.xlsx");
        }
    }
}
