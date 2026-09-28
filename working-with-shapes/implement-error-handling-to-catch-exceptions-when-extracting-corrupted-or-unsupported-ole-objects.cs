// Title: Extract embedded OLE objects from an Excel workbook in C# using Aspose.Cells with error handling for corrupted or unsupported objects
// AI Prompts: Write C# code that loads an Excel file with Aspose.Cells, iterates through each worksheet, extracts every embedded OLE object to a .bin file, and wraps the extraction in try‑catch blocks that capture CellsException and generic exceptions. | Update an existing Aspose.Cells OLE extraction routine to log the name of each OLE object and record detailed error messages when the object cannot be saved because it is corrupted or not supported. | Create a console application that validates the input path, ensures the output directory exists, and safely extracts OLE objects while handling missing files, workbook load failures, and unsupported OLE formats.
// Common Searches: c# aspnet extract ole objects from excel file with aspose.cells and handle errors | how to catch CellsException when extracting embedded OLE objects using Aspose.Cells | save embedded OLE objects as binary files from .xlsx using Aspose.Cells .NET | extract ole objects from multiple worksheets asp.net aspose.cells error handling | asp.net core extract ole objects from excel and log corrupted object warnings
// Tags: Aspose.Cells OLE object extraction | C# extract embedded OLE from Excel | handle corrupted OLE with CellsException | binary export of OLE objects .NET | reflection based OLE data retrieval Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// Demonstrates how to load an Excel workbook with Aspose.Cells, iterate through all worksheets, extract each embedded OLE object to a binary file, and use try‑catch blocks (including CellsException) to gracefully handle missing files, load failures, corrupted or unsupported OLE objects, and other unexpected errors.
class OleExtractor
{
    /// <param name="inputFile">Path to the source Excel file.</param>
    /// <param name="outputDir">Directory where extracted OLE objects will be saved.</param>
    public static void ExtractOleObjects(string inputFile, string outputDir)
    {
        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputFile))
        {
            Console.WriteLine($"Input file not found: {inputFile}");
            return;
        }

        Workbook workbook;
        try
        {
            // Load the workbook
            workbook = new Workbook(inputFile);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to load workbook '{inputFile}': {ex.Message}");
            return;
        }

        // Ensure the output directory exists
        Directory.CreateDirectory(outputDir);

        // Iterate through each worksheet in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Iterate through each OLE object in the worksheet
            foreach (OleObject ole in sheet.OleObjects)
            {
                // Build a unique file name for the extracted OLE object
                string fileName = $"{Path.GetFileNameWithoutExtension(inputFile)}_{sheet.Name}_{ole.Name}.bin";
                string outPath = Path.Combine(outputDir, fileName);

                try
                {
                    // Attempt to obtain the OLE object data via reflection (covers different API versions)
                    var oleDataProp = ole.GetType().GetProperty("OleObjectData") ??
                                      ole.GetType().GetProperty("ObjectData");

                    if (oleDataProp != null)
                    {
                        dynamic oleData = oleDataProp.GetValue(ole, null);
                        if (oleData != null)
                        {
                            // Save the OLE object data to a file
                            oleData.Save(outPath);
                            Console.WriteLine($"Extracted OLE object '{ole.Name}' to '{outPath}'.");
                            continue;
                        }
                    }

                    // If we reach here, no extractable data was found
                    Console.WriteLine($"OLE object '{ole.Name}' does not contain extractable data.");
                }
                catch (CellsException ex)
                {
                    // Handle Aspose.Cells specific errors (e.g., corrupted or unsupported OLE)
                    Console.WriteLine($"Failed to extract OLE object '{ole.Name}': {ex.Message}");
                }
                catch (Exception ex)
                {
                    // Handle any other unexpected errors
                    Console.WriteLine($"Unexpected error while extracting OLE object '{ole.Name}': {ex.Message}");
                }
            }
        }
    }

    // Entry point for the console application
    static void Main(string[] args)
    {
        try
        {
            if (args.Length < 2)
            {
                Console.WriteLine("Usage: OleExtractor <inputExcelFile> <outputDirectory>");
                return;
            }

            string inputFile = args[0];
            string outputDir = args[1];

            ExtractOleObjects(inputFile, outputDir);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Fatal error: {ex.Message}");
        }
    }
}
