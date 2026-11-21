// Title: Extract all embedded OLE objects from an Excel workbook to a folder with Aspose.Cells for .NET (C#)
// Common Searches: how to save embedded OLE objects from an .xlsx file using Aspose.Cells C# | Aspose.Cells C# iterate through OleObjects collection and export files | C# code to extract OLE objects from each worksheet in a workbook | batch export OLE objects to disk with Aspose.Cells for .NET | determine file extension when exporting OLE objects using Aspose.Cells
// Tags: Aspose.Cells export OLE objects to files | C# iterate worksheet OleObjects collection | save embedded OLE binary data from Excel | write OLE object data with Aspose.Cells | fallback .bin extension for unknown OLE format

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing; // Required for OleObject handling

// The example loads an Excel workbook with Aspose.Cells, creates an output directory, and loops through every worksheet and its OleObjects collection. For each OLE object it reads the raw ObjectData byte array, builds a unique filename using the workbook name, sheet name, object name, and a counter, assigns a default .bin extension (or custom logic), writes the data to disk, and tracks the number of successful extractions while handling errors.
class OleObjectExtractor
{
    static void Main()
    {
        // Path to the source workbook
        string workbookPath = "input.xlsx";

        // Folder where extracted OLE objects will be saved
        string outputFolder = "OleObjects";

        try
        {
            // Verify that the input file exists
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"Error: Workbook file \"{workbookPath}\" not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(workbookPath);

            // Ensure the output folder exists
            if (!Directory.Exists(outputFolder))
                Directory.CreateDirectory(outputFolder);

            int oleCounter = 0;

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all OLE objects in the current worksheet
                foreach (OleObject ole in sheet.OleObjects)
                {
                    try
                    {
                        // Retrieve the raw data of the OLE object
                        byte[] oleData = ole.ObjectData;

                        // Build a unique file name using workbook, sheet and OLE object names
                        string baseFileName = $"{Path.GetFileNameWithoutExtension(workbookPath)}_{sheet.Name}_{ole.Name}_{oleCounter}";

                        // Determine a suitable file extension (fallback to .bin)
                        string extension = GetExtensionFromOleObject(ole);

                        // Full path for the extracted file
                        string outputPath = Path.Combine(outputFolder, baseFileName + extension);

                        // Write the OLE object data to the file
                        File.WriteAllBytes(outputPath, oleData);

                        oleCounter++;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to extract OLE object \"{ole.Name}\" from sheet \"{sheet.Name}\": {ex.Message}");
                    }
                }
            }

            Console.WriteLine($"Extraction complete. {oleCounter} OLE object(s) saved to \"{outputFolder}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Helper method to determine a file extension for an OLE object.
    // Aspose.Cells does not expose the original format directly, so we default to .bin.
    private static string GetExtensionFromOleObject(OleObject ole)
    {
        // If needed, custom logic can be added here to infer the extension from the OLE object's data.
        return ".bin";
    }
}
