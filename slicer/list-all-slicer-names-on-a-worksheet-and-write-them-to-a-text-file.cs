// Title: List all slicer names in an Excel worksheet and export to a text file – Aspose.Cells for .NET
// Description: Loads an Excel workbook, accesses the first worksheet, iterates through its Slicers collection, captures each slicer's Name property, and writes the names line‑by‑line to a file named SlicerNames.txt. Includes handling for worksheets without slicers.
// Keywords: Aspose.Cells slicer enumeration | C# extract slicer names | export slicer list to txt | Excel slicer names Aspose | list slicers worksheet .NET
// Common Searches: how to get slicer names with Aspose.Cells C# | save Excel slicer names to a text file | list all slicers in a workbook using Aspose | C# code to export slicer names from worksheet
// Developer Intent: Obtain every slicer name from a worksheet and write the collection to a plain‑text file.
// Use Cases: Generate a quick inventory of slicers for documentation or audit purposes. | Validate naming conventions of slicers before distributing the workbook. | Feed slicer names into automated scripts that rename or relocate slicers across multiple files.
// AI Prompts: Write C# code with Aspose.Cells that lists slicer names and their captions, then saves them as CSV. | Explain how to gracefully handle a worksheet that contains no slicers when exporting names. | Show how to customize the output path and encoding when writing slicer names to a file.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Slicers;

// Loads an Excel workbook, accesses the first worksheet, iterates through its Slicers collection, captures each slicer's Name property, and writes the names line‑by‑line to a file named SlicerNames.txt. Includes handling for worksheets without slicers.
class Program
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");
        Worksheet worksheet = workbook.Worksheets[0];

        // Gather all slicer names from the worksheet
        List<string> slicerNames = new List<string>();
        foreach (Slicer slicer in worksheet.Slicers)
        {
            slicerNames.Add(slicer.Name);
        }

        // Write the slicer names to a text file
        File.WriteAllLines("SlicerNames.txt", slicerNames);
    }
}
