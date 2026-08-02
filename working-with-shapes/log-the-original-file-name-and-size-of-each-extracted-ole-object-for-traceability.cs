// Title: C# – Log Original File Name and Size When Extracting OLE Objects with Aspose.Cells
// Description: Loads an Excel workbook, walks through each worksheet, extracts every OLE object, records its original file name (or a generated placeholder) and binary size, writes the details to the console, saves the object to a designated folder, and finally saves the workbook.
// Keywords: Aspose.Cells OLE extraction C# | log OLE object size | original file name OLE Excel | traceability embedded OLE | extract linked OLE objects | C# Excel OLE handling
// Common Searches: how to extract OLE objects from Excel using Aspose.Cells | log OLE object file name and size in C# | retrieve original file name of linked OLE in workbook | save extracted OLE objects with original names | Aspose.Cells OLE extraction example
// Developer Intent: The developer needs to capture the source name and byte size of each OLE object while extracting it from an Excel file for audit and traceability purposes.
// Use Cases: Generate an audit report of all OLE objects before archiving a workbook. | Enforce size limits on extracted OLE content during data migration. | Populate a processing folder with OLE files named after their original sources for downstream analysis.
// AI Prompts: Create C# code that extracts OLE objects with Aspose.Cells and outputs a CSV containing worksheet, index, original name, and size. | Suggest a strategy to assign unique identifiers to OLE objects that lack a source file name. | Explain how to embed this extraction and logging routine into a larger file‑processing pipeline with robust error handling and structured logging.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace OleObjectExtractionDemo
{
    // Loads an Excel workbook, walks through each worksheet, extracts every OLE object, records its original file name (or a generated placeholder) and binary size, writes the details to the console, saves the object to a designated folder, and finally saves the workbook.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source workbook containing OLE objects
            string inputPath = "InputWorkbook.xlsx";

            // Load the workbook (create/load rule)
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Get the collection of OLE objects in the current worksheet
                OleObjectCollection oleObjects = sheet.OleObjects;

                // Process each OLE object
                for (int i = 0; i < oleObjects.Count; i++)
                {
                    OleObject ole = oleObjects[i];

                    // Determine if the OLE object is linked or embedded
                    bool isLinked = ole.IsLink;

                    // Original file name (for linked objects) or generated name for embedded ones
                    string originalFileName = isLinked ? ole.ObjectSourceFullName : $"EmbeddedOle_{i + 1}.bin";

                    // Size of the OLE object data in bytes
                    int dataSize = ole.ObjectData?.Length ?? 0;

                    // Log the information
                    Console.WriteLine($"Worksheet: {sheet.Name}, OLE Index: {i}");
                    Console.WriteLine($"  Original File Name: {originalFileName}");
                    Console.WriteLine($"  Data Size (bytes): {dataSize}");

                    // Extract the OLE object to a file for traceability
                    // Use the original file name if available; otherwise, use the generated name
                    string outputFilePath = Path.Combine("ExtractedOleObjects", Path.GetFileName(originalFileName));

                    // Ensure the output directory exists
                    Directory.CreateDirectory(Path.GetDirectoryName(outputFilePath));

                    // Write the binary data to the file
                    if (ole.ObjectData != null && ole.ObjectData.Length > 0)
                    {
                        File.WriteAllBytes(outputFilePath, ole.ObjectData);
                        Console.WriteLine($"  Extracted to: {outputFilePath}");
                    }
                    else
                    {
                        Console.WriteLine("  No binary data to extract.");
                    }
                }
            }

            // Save the workbook after processing (save rule)
            workbook.Save("ProcessedWorkbook.xlsx");
        }
    }
}
