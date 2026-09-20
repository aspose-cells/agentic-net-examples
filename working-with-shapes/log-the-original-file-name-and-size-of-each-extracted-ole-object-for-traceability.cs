// Title: Log original OLE object file name and byte size while iterating worksheets using Aspose.Cells for .NET
// AI Prompts: Write C# code that uses Aspose.Cells to walk through every worksheet, read each embedded OLE object's Name property and ObjectData length, and output the worksheet name, OLE index, original name, and size in bytes to the console. | Modify the OLE extraction loop to also write each object's binary data to a temporary file named after its original OLE name, while still logging the name and size for traceability. | Add per‑worksheet try/catch blocks that capture exceptions when accessing OLE objects, continue processing remaining sheets, and log the error together with the worksheet and OLE index.
// Common Searches: how to enumerate OLE objects in an Excel file with Aspose.Cells C# | retrieve embedded OLE object name and size using Aspose.Cells .NET | log OLE object details while extracting from multiple worksheets in C# | Aspose.Cells get byte length of OLE object data in .xlsx | continue processing other sheets when OLE extraction fails Aspose.Cells
// Tags: enumerate OLE objects Aspose.Cells | log embedded OLE metadata .NET | extract OLE object size C# | retrieve OLE original file name Aspose.Cells | worksheet‑level OLE extraction error handling

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// The program loads an Excel workbook, iterates through each worksheet, and for every embedded OLE object obtains its original Name and binary data length. It logs the worksheet name, OLE index, original name, and size in bytes to the console, with per‑worksheet error handling to ensure processing continues even if an object fails to load.
class Program
{
    static void Main()
    {
        string inputPath = "input.xlsx";

        // Ensure the input file exists before attempting to load it
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                try
                {
                    // Get the collection of OLE objects on the current worksheet
                    OleObjectCollection oleObjects = sheet.OleObjects;

                    // Process each OLE object
                    for (int i = 0; i < oleObjects.Count; i++)
                    {
                        OleObject ole = oleObjects[i];

                        // Original name of the OLE object (if available)
                        string originalName = ole.Name;

                        // Raw binary data of the OLE object
                        byte[] data = ole.ObjectData; // Correct property to retrieve OLE data

                        // Size of the OLE object in bytes
                        long size = data != null ? data.Length : 0;

                        // Log the information for traceability
                        Console.WriteLine($"Worksheet: {sheet.Name}, OLE Index: {i}, Name: {originalName}, Size: {size} bytes");
                    }
                }
                catch (Exception exSheet)
                {
                    // Handle errors specific to a worksheet without stopping the whole process
                    Console.WriteLine($"Error processing worksheet '{sheet.Name}': {exSheet.Message}");
                }
            }
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
