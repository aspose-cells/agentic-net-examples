// Title: How to retrieve and print slicer names from an Excel worksheet using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# get list of slicer names from worksheet | iterate through slicers in an Excel file using Aspose.Cells .NET | read slicer collection via reflection in Aspose.Cells C# example | check if worksheet contains slicers with Aspose.Cells and output their names | handle missing slicer feature in Aspose.Cells when loading a workbook
// Tags: Aspose.Cells slicer enumeration C# | worksheet slicer collection retrieval | reflection access to slicer name | Excel slicer listing with Aspose.Cells | unsupported slicer feature handling

using System;
using System.IO;
using System.Collections;
using Aspose.Cells;

// The sample loads a workbook, accesses the first worksheet, uses reflection to obtain the Slicers collection, iterates over each slicer, reads its Name property via reflection, and writes each name to the console, while gracefully handling cases where no slicers exist or the slicer feature is unavailable.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file '{inputPath}' was not found.");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (by index)
            Worksheet worksheet = workbook.Worksheets[0];

            // Attempt to retrieve slicers via reflection (avoids compile‑time dependency on Slicer types)
            object slicersObj = worksheet.GetType().GetProperty("Slicers")?.GetValue(worksheet);

            if (slicersObj is IEnumerable slicersEnumerable)
            {
                bool any = false;
                foreach (var slicer in slicersEnumerable)
                {
                    // Use reflection to read the Name property of each slicer
                    var nameProp = slicer.GetType().GetProperty("Name");
                    var name = nameProp?.GetValue(slicer);
                    Console.WriteLine(name?.ToString() ?? "Unnamed Slicer");
                    any = true;
                }

                if (!any)
                {
                    Console.WriteLine("No slicers found in the worksheet.");
                }
            }
            else
            {
                Console.WriteLine("Slicer feature is not supported in this version of Aspose.Cells or no slicers exist.");
            }

            // If changes are made and need to be saved, uncomment the line below
            // workbook.Save("output.xlsx");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
