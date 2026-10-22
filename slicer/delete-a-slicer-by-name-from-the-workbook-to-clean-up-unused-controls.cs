// Title: How to delete a slicer by name from an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads an .xlsx file with Aspose.Cells, searches all worksheets for a slicer named "MySlicer", deletes it, and saves the workbook. | Show a safe pattern that checks the input file exists, iterates through Worksheet.Slicers, and removes a slicer by its Name property using Aspose.Cells. | Provide a try‑catch example that removes a specific slicer, logs any exceptions, and writes the updated workbook to a new file.
// Common Searches: asp.net delete slicer by name from Excel workbook using Aspose.Cells | C# program to remove specific slicer from .xlsx file | how to iterate worksheet slicers collection Aspose.Cells | remove unused slicer in Excel with Aspose.Cells C# example
// Tags: Aspose.Cells delete slicer by name | C# remove Excel slicer programmatically | Worksheet slicer collection manipulation Aspose.Cells | clean up unused slicers .NET | save workbook after slicer removal Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads InputWorkbook.xlsx, verifies the file exists, iterates through each worksheet's Slicers collection to find a slicer named "MySlicer", removes it, and saves the modified workbook as OutputWorkbook.xlsx while handling potential exceptions.
class DeleteSlicerExample
{
    static void Main()
    {
        // Paths for input and output workbooks
        string inputFile = "InputWorkbook.xlsx";
        string outputFile = "OutputWorkbook.xlsx";
        // Name of the slicer to delete
        string slicerName = "MySlicer";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputFile))
        {
            Console.WriteLine($"Input file not found: {inputFile}");
            return;
        }

        try
        {
            // Load the workbook (lifecycle rule: load)
            Workbook workbook = new Workbook(inputFile);

            // Iterate through worksheets to locate the slicer
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Access the slicer collection of the current worksheet
                var slicers = sheet.Slicers;

                // Find the index of the slicer with the specified name
                int slicerIndex = -1;
                for (int i = 0; i < slicers.Count; i++)
                {
                    var s = slicers[i];
                    if (s.Name.Equals(slicerName, StringComparison.OrdinalIgnoreCase))
                    {
                        slicerIndex = i;
                        break;
                    }
                }

                // If found, remove the slicer and exit the loop
                if (slicerIndex >= 0)
                {
                    slicers.RemoveAt(slicerIndex);
                    break; // Assuming slicer names are unique
                }
            }

            // Save the modified workbook (lifecycle rule: save)
            workbook.Save(outputFile);
            Console.WriteLine($"Workbook saved successfully to {outputFile}");
        }
        catch (Exception ex)
        {
            // Handle any runtime exceptions gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
