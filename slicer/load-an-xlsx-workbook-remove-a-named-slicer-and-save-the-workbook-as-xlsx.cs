// Title: C# – Remove a Named Slicer from an XLSX Workbook using Aspose.Cells and Save
// Description: Loads an existing XLSX file with Aspose.Cells, finds a slicer named "MySlicer" in the first worksheet, removes it from the SlicerCollection, and saves the updated workbook as a new XLSX document.
// Keywords: Aspose.Cells | C# | .NET | remove slicer | delete Excel slicer | XLSX workbook | SlicerCollection | programmatic Excel | sample code | GitHub example
// Common Searches: how to delete a slicer in Excel using Aspose.Cells C# | remove named slicer from XLSX with Aspose.Cells .NET | Aspose.Cells code to delete slicer programmatically | C# example for removing Excel slicer and saving workbook | Aspose.Cells slicer removal sample on GitHub
// Developer Intent: Programmatically delete a specific slicer from an existing XLSX workbook and write the changes back to disk.
// Use Cases: Strip temporary slicers before distributing a report. | Automate cleanup of dynamically added slicers in batch‑processed workbooks. | Prepare a clean version of a dashboard by removing obsolete slicer controls.
// AI Prompts: Generate C# code that uses Aspose.Cells to locate and remove a slicer named "MySlicer" from an XLSX file. | Explain how to safely handle the situation when the specified slicer does not exist in the workbook. | Show a step‑by‑step example of loading a workbook, iterating the SlicerCollection, deleting a slicer, and saving the file.

using System;
using Aspose.Cells;
using Aspose.Cells.Slicers;

// Loads an existing XLSX file with Aspose.Cells, finds a slicer named "MySlicer" in the first worksheet, removes it from the SlicerCollection, and saves the updated workbook as a new XLSX document.
class RemoveSlicerDemo
{
    static void Main()
    {
        // Load the existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Access the worksheet that contains the slicer (adjust index if needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Get the slicer collection from the worksheet
        SlicerCollection slicers = worksheet.Slicers;

        // Name of the slicer to be removed
        string slicerName = "MySlicer";

        // Locate the slicer by its name
        Slicer slicerToRemove = null;
        foreach (Slicer slicer in slicers)
        {
            if (slicer.Name == slicerName)
            {
                slicerToRemove = slicer;
                break;
            }
        }

        // Remove the slicer if it was found
        if (slicerToRemove != null)
        {
            slicers.Remove(slicerToRemove);
        }

        // Save the modified workbook as XLSX
        workbook.Save("output.xlsx");
    }
}
