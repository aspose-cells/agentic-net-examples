// Title: Aspose.Cells for .NET – Get and List All Slicer Names in a Worksheet (C#)
// Description: A concise C# sample that shows how to create or load a workbook with Aspose.Cells, access the first worksheet, obtain its SlicerCollection, iterate through each Slicer, write the slicer’s Name to the console, and save the workbook. The code covers the full workbook lifecycle while focusing on slicer enumeration.
// Keywords: Aspose.Cells slicer collection C# | list slicer names .NET | enumerate worksheet slicers | Aspose.Cells Slicer API | C# get slicer objects | Aspose.Cells workbook slicer iteration
// Common Searches: C# loop through slicers Aspose.Cells | How to list slicer names in a worksheet using Aspose.Cells | Aspose.Cells get all slicers from a sheet | Retrieve slicer objects with Aspose.Cells .NET
// Developer Intent: Extract every slicer on a sheet and display its identifier.
// Use Cases: Validate that required slicers exist before generating a report. | Log slicer names for auditing workbook structure. | Collect slicer identifiers to programmatically modify their properties later.
// AI Prompts: Show how to rename each slicer in the collection with Aspose.Cells for .NET. | Provide code to check for a slicer by name and delete it if found. | Explain how to add a new slicer to a worksheet and configure its initial settings.

using System;
using Aspose.Cells;
using Aspose.Cells.Slicers;

// A concise C# sample that shows how to create or load a workbook with Aspose.Cells, access the first worksheet, obtain its SlicerCollection, iterate through each Slicer, write the slicer’s Name to the console, and save the workbook. The code covers the full workbook lifecycle while focusing on slicer enumeration.
class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook(); // lifecycle: create

        // Get the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Retrieve the slicer collection from the worksheet
        SlicerCollection slicers = worksheet.Slicers;

        // Iterate through each slicer and log its name
        foreach (Slicer slicer in slicers)
        {
            Console.WriteLine("Slicer Name: " + slicer.Name);
        }

        // Save the workbook (lifecycle: save)
        workbook.Save("output.xlsx");
    }
}
