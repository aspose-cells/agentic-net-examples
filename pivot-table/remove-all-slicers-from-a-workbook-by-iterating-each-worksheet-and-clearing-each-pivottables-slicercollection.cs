// Title: Remove All Slicers from an Excel Workbook with Aspose.Cells for .NET
// Description: Loads a workbook, iterates each worksheet, clears the SlicerCollection to delete every slicer, and saves the file without slicers.
// Keywords: Aspose.Cells remove slicers | C# clear slicer collection | delete all slicers Excel | SlicerCollection.Clear Aspose | programmatic slicer removal | Excel workbook cleanup .NET
// Common Searches: how to delete all slicers using Aspose.Cells C# | clear slicer collection for each worksheet Aspose | remove slicers from Excel file programmatically | Aspose.Cells example to strip slicers
// Developer Intent: Eliminate every slicer in a workbook by clearing each worksheet’s SlicerCollection.
// Use Cases: Prepare a clean template before distribution by stripping slicers. | Automate batch processing of workbooks to remove slicers in a data‑export pipeline. | Clean up after pivot‑table updates when slicers are no longer required.
// AI Prompts: Generate C# code with Aspose.Cells that removes all slicers from a workbook and saves the result. | Explain the performance impact of SlicerCollection.Clear() on large Excel files. | Show how to target specific worksheets for slicer removal using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Slicers;

// Loads a workbook, iterates each worksheet, clears the SlicerCollection to delete every slicer, and saves the file without slicers.
class RemoveAllSlicers
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("InputWorkbook.xlsx");

        // Iterate through each worksheet in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Get the slicer collection for the current worksheet
            SlicerCollection slicers = sheet.Slicers;

            // Clear all slicers from this worksheet
            slicers.Clear();
        }

        // Save the modified workbook (replace with your desired output path)
        workbook.Save("WorkbookWithoutSlicers.xlsx");
    }
}
