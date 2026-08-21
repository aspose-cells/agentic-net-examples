// Title: C# – Extract and Safely Handle Corrupted or Unsupported OLE Objects with Aspose.Cells
// Description: Demonstrates how to load an Excel workbook, iterate through each worksheet's OLE objects, extract the embedded data to files, and apply robust exception handling for corrupted, unsupported, or stream‑related OLE objects using Aspose.Cells for .NET. The sample logs errors, skips failing objects, and ensures the workbook is saved after processing.
// Keywords: Aspose.Cells OLE extraction | C# OLE object handling | catch CellsException | unsupported OLE data | corrupted OLE object | exception handling Aspose.Cells | extract embedded OLE from Excel | Aspose.Cells .NET error handling
// Common Searches: Aspose.Cells extract OLE object C# | how to handle corrupted OLE objects in Excel | catch CellsException for unsupported OLE data | C# code to save embedded OLE files from workbook | error handling when reading OLE objects with Aspose
// Developer Intent: Add comprehensive try‑catch blocks to extract OLE objects while gracefully handling corrupted or unsupported data without aborting the whole operation.
// Use Cases: Bulk extraction of all embedded OLE objects from a workbook, saving each to a uniquely named file. | Logging detailed CellsException information for objects that cannot be read, enabling later analysis. | Continuing worksheet processing after an extraction failure, ensuring the final workbook is still saved.
// AI Prompts: Write a reusable C# method that extracts OLE objects with Aspose.Cells and returns a status report for each object. | Create a PowerShell script that calls the provided C# code to batch‑process multiple Excel files and aggregate extraction logs. | Suggest how to implement retry logic for transient OLE extraction errors while preserving existing CellsException handling.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to load an Excel workbook, iterate through each worksheet's OLE objects, extract the embedded data to files, and apply robust exception handling for corrupted, unsupported, or stream‑related OLE objects using Aspose.Cells for .NET. The sample logs errors, skips failing objects, and ensures the workbook is saved after processing.
class ExtractOleObjects
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through each worksheet
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Iterate through each OLE object in the worksheet
            for (int i = 0; i < sheet.OleObjects.Count; i++)
            {
                OleObject ole = sheet.OleObjects[i];
                try
                {
                    // Attempt to retrieve the embedded OLE data.
                    // This may throw if the object is corrupted or unsupported.
                    byte[] data = ole.ObjectData;

                    if (data != null && data.Length > 0)
                    {
                        // Determine a file name for the extracted data.
                        string extension = Path.GetExtension(ole.ObjectSourceFullName);
                        if (string.IsNullOrEmpty(extension))
                            extension = ".bin";

                        string outPath = $"OleObject_{sheet.Name}_{i}{extension}";

                        // Write the extracted bytes to disk.
                        File.WriteAllBytes(outPath, data);
                        Console.WriteLine($"Extracted OLE object to {outPath}");
                    }
                }
                // Handle known Aspose.Cells exceptions related to corrupted or unsupported OLE data.
                catch (CellsException ex) when (ex.Code == ExceptionType.FileCorrupted ||
                                               ex.Code == ExceptionType.UnsupportedFeature ||
                                               ex.Code == ExceptionType.UnsupportedStream)
                {
                    Console.WriteLine($"Failed to extract OLE object at index {i} on sheet '{sheet.Name}': {ex.Message} (Code: {ex.Code})");
                }
                // Fallback for any other unexpected errors.
                catch (Exception ex)
                {
                    Console.WriteLine($"Unexpected error extracting OLE object at index {i} on sheet '{sheet.Name}': {ex.Message}");
                }
            }
        }

        // Optionally save the workbook after processing.
        workbook.Save("output.xlsx");
    }
}
