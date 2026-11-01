// Title: Extract embedded OLE objects from an Excel workbook and log their original file names and sizes using Aspose.Cells for .NET
// AI Prompts: Generate C# code that opens an Excel file with Aspose.Cells, walks through every worksheet, extracts each OleObject's binary data, saves it to a designated folder with an inferred file extension, and prints the original object name together with the byte count. | Update an existing OleObject extraction loop to output the workbook name, sheet name, OLE index, original file name, and extracted data size to the console for traceability. | Create a reusable method that takes a workbook path and output directory, extracts all embedded OLE objects, writes them to disk, and returns a list containing OriginalName, SavedPath, and SizeInBytes.
// Common Searches: c# aspose.cells extract ole objects from excel and retrieve original filename | how to log size of extracted OLE objects using Aspose.Cells .NET | save embedded OLE objects with inferred extension Aspose.Cells example | traceability of OLE extraction from workbook using Aspose.Cells | retrieve OleObject.ObjectData and write to file in C#
// Tags: Aspose.Cells extract OLE objects C# | log extracted OLE object size .NET | save embedded OLE binary with inferred extension | OleObject.ObjectData retrieval Aspose.Cells | traceability of Excel OLE extraction | extract OLE objects to folder Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing; // Required for OleObject and OleObjectCollection

// The example loads an Excel workbook, iterates all worksheets, extracts each embedded OLE object, infers a file extension, saves the binary to a folder, and writes the original object name and byte size to the console for traceability.
class Program
{
    static void Main()
    {
        // Path to the source Excel file
        string workbookPath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(workbookPath))
        {
            Console.WriteLine($"Error: The file '{workbookPath}' was not found.");
            return;
        }

        // Directory where extracted OLE objects will be saved
        string outputDir = "ExtractedOleObjects";

        // Ensure the output directory exists
        Directory.CreateDirectory(outputDir);

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(workbookPath);

            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Get the collection of OLE objects on the current sheet
                OleObjectCollection oleObjects = sheet.OleObjects;

                // Process each OLE object
                for (int i = 0; i < oleObjects.Count; i++)
                {
                    OleObject ole = oleObjects[i];

                    // Retrieve the binary data of the OLE object using the correct API
                    byte[] data;
                    try
                    {
                        data = ole.ObjectData; // Correct property for .NET
                        if (data == null || data.Length == 0)
                        {
                            Console.WriteLine($"No data found for OLE object '{ole.Name}' on sheet '{sheet.Name}'.");
                            continue;
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to get data for OLE object '{ole.Name}' on sheet '{sheet.Name}': {ex.Message}");
                        continue;
                    }

                    // Attempt to infer a file extension from the OLE object's name
                    string extension = ".bin";
                    string originalName = ole.Name; // May contain the original file name
                    if (!string.IsNullOrEmpty(originalName))
                    {
                        string extFromName = Path.GetExtension(originalName);
                        if (!string.IsNullOrEmpty(extFromName))
                            extension = extFromName;
                    }

                    // Build a unique file name for the extracted object
                    string fileName = $"{Path.GetFileNameWithoutExtension(workbookPath)}_Sheet{sheet.Index}_Ole{i + 1}{extension}";
                    string filePath = Path.Combine(outputDir, fileName);

                    // Save the OLE object to disk
                    try
                    {
                        File.WriteAllBytes(filePath, data);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to write OLE object to file '{filePath}': {ex.Message}");
                        continue;
                    }

                    // Log the original name and size for traceability
                    Console.WriteLine($"Extracted OLE object from sheet '{sheet.Name}' (Index {i + 1}):");
                    Console.WriteLine($"  Original Name: {originalName}");
                    Console.WriteLine($"  Saved As: {fileName}");
                    Console.WriteLine($"  Size: {data.Length} bytes");
                }
            }

            Console.WriteLine("OLE object extraction completed.");
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
