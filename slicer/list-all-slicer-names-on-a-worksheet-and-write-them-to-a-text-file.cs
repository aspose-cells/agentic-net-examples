// Title: List all slicer names in an Excel worksheet and export them to a text file with Aspose.Cells for .NET
// AI Prompts: Generate C# code using Aspose.Cells that loads a workbook, iterates over the worksheet's SlicerCollection, and writes each slicer’s Name to a .txt file. | Show how to retrieve slicer names from a specific worksheet and save the list as plain‑text using the Aspose.Cells .NET API. | Provide a complete example that opens input.xlsx, extracts all slicer identifiers, and creates SlicerNames.txt with one name per line.
// Common Searches: Aspose.Cells C# get names of all slicers on a worksheet | How to export Excel slicer names to a text file using .NET | C# code to list slicer objects in a workbook with Aspose.Cells | Save slicer identifiers from an Excel file to a .txt file in C# | Retrieve slicer collection and write names to file Aspose.Cells example
// Tags: Aspose.Cells SlicerCollection enumeration | Aspose.Cells generate text file of slicer names | list worksheet slicers Aspose.Cells | C# write Excel slicer names to file | retrieve slicer Name property Aspose

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Slicers;

// The program loads an existing workbook, accesses the first worksheet, iterates through its SlicerCollection to collect each slicer's Name, and writes the names line‑by‑line to a text file named SlicerNames.txt.
class ListSlicerNames
{
    static void Main()
    {
        // Load an existing workbook that contains slicers
        Workbook workbook = new Workbook("input.xlsx");

        // Get the first worksheet (or specify the desired worksheet index/name)
        Worksheet worksheet = workbook.Worksheets[0];

        // Retrieve the slicer collection from the worksheet
        SlicerCollection slicers = worksheet.Slicers;

        // Prepare a list to hold slicer names
        List<string> slicerNames = new List<string>();

        // Iterate through each slicer and collect its Name property
        foreach (Slicer slicer in slicers)
        {
            slicerNames.Add(slicer.Name);
        }

        // Write all slicer names to a text file, one name per line
        File.WriteAllLines("SlicerNames.txt", slicerNames);
    }
}
