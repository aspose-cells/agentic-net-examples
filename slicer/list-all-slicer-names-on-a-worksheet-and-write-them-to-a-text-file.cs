using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Slicers;

class ListSlicerNames
{
    static void Main()
    {
        // Load an existing workbook that contains slicers
        Workbook workbook = new Workbook("input.xlsx");

        // Get the first worksheet (or specify the desired one)
        Worksheet worksheet = workbook.Worksheets[0];

        // Retrieve the slicer collection from the worksheet
        SlicerCollection slicers = worksheet.Slicers;

        // Prepare a list to hold slicer names
        List<string> slicerNames = new List<string>();

        // Iterate through all slicers and collect their names
        for (int i = 0; i < slicers.Count; i++)
        {
            Slicer slicer = slicers[i];
            slicerNames.Add(slicer.Name);
        }

        // Write the slicer names to a text file, one name per line
        File.WriteAllLines("SlicerNames.txt", slicerNames);

        // Optionally save the workbook (demonstrating the save rule)
        workbook.Save("output.xlsx");
    }
}