using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Slicers;

class ListSlicerNames
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Get the first worksheet (or specify the desired one)
        Worksheet worksheet = workbook.Worksheets[0];

        // Access the slicer collection on the worksheet
        SlicerCollection slicers = worksheet.Slicers;

        // Collect slicer names
        List<string> slicerNames = new List<string>();
        for (int i = 0; i < slicers.Count; i++)
        {
            slicerNames.Add(slicers[i].Name);
        }

        // Write the slicer names to a text file
        File.WriteAllLines("SlicerNames.txt", slicerNames);

        // Optionally save the workbook if any changes were made
        workbook.Save("output.xlsx");
    }
}