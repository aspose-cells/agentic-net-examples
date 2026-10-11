// Title: C# – Extract embedded OLE objects from an Excel file using Aspose.Cells with comprehensive exception handling
// AI Prompts: Generate C# code that loads an Excel workbook with Aspose.Cells, verifies the file exists, iterates all worksheets, and extracts each OleObject's raw data to a uniquely named .bin file, wrapping the extraction in try‑catch blocks to log failures without stopping the loop. | Create a reusable C# method that accepts a workbook path and an output folder, returns a list of successfully saved OLE object file paths, and captures any exceptions thrown during workbook loading or OleObject extraction, including corrupted or unsupported objects. | Write a C# console program that uses Aspose.Cells to enumerate OleObjects, writes their ObjectData to disk, and implements detailed error handling for FileNotFoundException, workbook load errors, and per‑object extraction errors, outputting concise console messages.
// Common Searches: aspnet extract ole objects from excel with aspose.cells and handle extraction errors | c# code to save embedded ole objects as binary files using Aspose.Cells | how to catch exceptions for corrupted OLE objects when using Aspose.Cells | extract all OLE objects from each worksheet in an Excel workbook using Aspose.Cells .NET | log failed OLE object extraction without aborting the process in C#
// Tags: Aspose.Cells extract OLE objects with exception handling | C# save embedded OLE data to binary files | handle corrupted OLE objects in Excel extraction | validate workbook existence before Aspose.Cells load | iterate worksheets OleObject extraction .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing; // Required for OleObject

// The example loads an Excel workbook, ensures the input file exists, creates an output directory, and iterates through every worksheet and its OleObjects. Each object's raw data is written to a uniquely named .bin file inside a try‑catch block, so corrupted or unsupported OLE objects are reported but do not stop the overall extraction process.
class OleExtractionWithErrorHandling
{
    static void Main()
    {
        // Path to the input workbook
        string inputPath = "InputWorkbook.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        Workbook workbook;
        try
        {
            // Load the workbook
            workbook = new Workbook(inputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to load workbook: {ex.Message}");
            return;
        }

        // Directory where extracted OLE objects will be saved
        string outputDir = "ExtractedOleObjects";
        Directory.CreateDirectory(outputDir);

        // Base name for generated files (derived from input file name)
        string baseFileName = Path.GetFileNameWithoutExtension(inputPath);

        // Iterate through all worksheets
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Iterate through all OLE objects in the current worksheet using index
            for (int i = 0; i < sheet.OleObjects.Count; i++)
            {
                OleObject ole = sheet.OleObjects[i];

                // Build a unique file name for the extracted OLE object
                string fileName = $"{baseFileName}_{sheet.Name}_Ole_{i}.bin";
                string outputPath = Path.Combine(outputDir, fileName);

                try
                {
                    // Retrieve the raw OLE object data
                    byte[] data = ole.ObjectData;

                    // Write the data to a file
                    File.WriteAllBytes(outputPath, data);
                    Console.WriteLine($"Successfully extracted OLE object to: {outputPath}");
                }
                catch (Exception ex)
                {
                    // Handle extraction errors (corrupted or unsupported OLE objects)
                    Console.WriteLine($"Failed to extract OLE object (Index: {i}) in worksheet '{sheet.Name}': {ex.Message}");
                }
            }
        }

        // No modifications were made, so saving the workbook is optional.
        // workbook.Save("OutputWorkbook.xlsx");
    }
}
