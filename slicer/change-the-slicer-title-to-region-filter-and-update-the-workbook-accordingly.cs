// Title: Set slicer title to "Region Filter" across an Excel workbook with Aspose.Cells for .NET
// Description: Loads a workbook, iterates through each worksheet’s SlicerCollection, assigns the (obsolete but functional) Slicer.Title property the value "Region Filter", and saves the updated file. Shows how to rename slicer captions in bulk using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | Excel slicer title | Slicer.Title | update slicer caption | bulk slicer rename | programmatic Excel filter label
// Common Searches: Aspose.Cells change slicer title | C# set slicer caption Excel | update all slicer titles programmatically | rename slicer label Aspose.Cells | bulk edit slicer titles .NET
// Developer Intent: Rename every slicer in a workbook to a single, consistent title ('Region Filter') and save the modified workbook.
// Use Cases: Standardize slicer headings across multiple worksheets before distributing a report. | Automate caption updates in a nightly batch that regenerates Excel dashboards. | Prepare a template workbook so end‑users see a uniform filter description.
// AI Prompts: Show C# code that sets all slicer titles to a custom string using Aspose.Cells. | Give an example that safely handles workbooks with no slicers while renaming titles. | Explain the deprecation of Slicer.Title and suggest the recommended approach for updating slicer captions.

using System;
using Aspose.Cells;
using Aspose.Cells.Slicers;

// Loads a workbook, iterates through each worksheet’s SlicerCollection, assigns the (obsolete but functional) Slicer.Title property the value "Region Filter", and saves the updated file. Shows how to rename slicer captions in bulk using Aspose.Cells for .NET.
class UpdateSlicerTitle
{
    static void Main()
    {
        // Load the existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through all worksheets
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Access the slicer collection of the current worksheet
            SlicerCollection slicers = sheet.Slicers;

            // Loop through each slicer and set its title to "Region Filter"
            for (int i = 0; i < slicers.Count; i++)
            {
                Slicer slicer = slicers[i];
                // The Title property is obsolete but still functional for setting the slicer title
                slicer.Title = "Region Filter";
            }
        }

        // Save the updated workbook (replace with your desired output path)
        workbook.Save("output.xlsx");
    }
}
