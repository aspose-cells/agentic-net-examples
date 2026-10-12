// Title: Extract every embedded OLE object from an Excel workbook to separate files using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that loops through all worksheets, accesses each OleObject, and writes its ObjectData to a file named after the object in a given output directory. | Enhance the extraction routine to infer the original file extension of each OLE object from its binary stream and save the file with that extension instead of a generic .bin file. | Add comprehensive error handling and create a log file that records workbook name, worksheet, OLE object name, extraction path, and any exceptions that occur.
// Common Searches: c# aspocells how to export embedded ole objects from excel to a folder | aspocells iterate oleobjects in all worksheets and save each to a file | extract ole object binary data from an Excel workbook using aspocells .net | save embedded excel objects as separate files with aspocells c# example
// Tags: aspocells ole object extraction | c# write ole object binary | iterate worksheets oleobjects aspocells | save embedded excel objects c# | batch ole extraction aspocells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample iterates through every worksheet in a workbook, extracts each OleObject's binary data, and writes it to a file in a specified output folder, handling missing data and logging extraction results.
class OleObjectExtractor
{
    static void Main()
    {
        // Path to the source workbook
        string workbookPath = @"C:\Path\To\Your\Workbook.xlsx";

        // Folder where extracted OLE objects will be saved
        string outputFolder = @"C:\Path\To\ExtractedOleObjects";

        try
        {
            // Ensure the output folder exists
            if (!Directory.Exists(outputFolder))
            {
                Directory.CreateDirectory(outputFolder);
            }

            // Verify the workbook file exists before loading
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"Error: Workbook file not found at '{workbookPath}'.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(workbookPath);

            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through each OLE object on the worksheet
                foreach (OleObject oleObject in sheet.OleObjects)
                {
                    try
                    {
                        // Retrieve the raw binary data of the OLE object
                        byte[] oleData = oleObject.ObjectData;

                        if (oleData == null || oleData.Length == 0)
                        {
                            Console.WriteLine($"Warning: OLE object '{oleObject.Name}' contains no data.");
                            continue;
                        }

                        // Build a file name for the extracted object.
                        // Using the OLE object's name and adding a generic .bin extension.
                        string fileName = $"{oleObject.Name}.bin";

                        // Full path for the extracted file
                        string outputPath = Path.Combine(outputFolder, fileName);

                        // Write the binary data to the file system
                        File.WriteAllBytes(outputPath, oleData);

                        // Output a message indicating success
                        Console.WriteLine($"Extracted OLE object '{oleObject.Name}' to '{outputPath}'.");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to extract OLE object '{oleObject.Name}': {ex.Message}");
                    }
                }
            }

            Console.WriteLine("Extraction complete.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
