// Title: C# – Extract all embedded OLE objects from an Excel workbook with Aspose.Cells
// Description: Loads an Excel file using Aspose.Cells for .NET, creates an output directory, walks through every worksheet and its OleObjectCollection, retrieves each object's binary data (ObjectData or FullObjectBin), determines a suitable file extension, and saves the content as uniquely named files in the target folder.
// Keywords: Aspose.Cells | C# | .NET | OLE object extraction | Excel embedded files | OleObjectCollection | ObjectData | FullObjectBin | save to folder | binary data export
// Common Searches: Aspose.Cells extract OLE objects C# | How to save embedded Excel objects with .NET | Export OleObjectCollection to files | Retrieve binary data of OLE objects in Aspose.Cells | C# code to extract embedded Word/PDF from Excel
// Developer Intent: Export every embedded OLE object from an Excel workbook to individual files on disk.
// Use Cases: Archive embedded documents (Word, PDF, images) from a financial report before sharing. | Migrate chart or diagram OLE objects to a content‑management system as separate files. | Create a backup of all embedded assets prior to batch‑processing or workbook modification.
// AI Prompts: Write C# code that uses Aspose.Cells to extract all OLE objects from a workbook while preserving original filenames when available. | Provide a method that returns a list of file paths for the extracted OLE objects and handles missing source names gracefully. | Explain the fallback from ObjectData to FullObjectBin when extracting embedded objects with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Loads an Excel file using Aspose.Cells for .NET, creates an output directory, walks through every worksheet and its OleObjectCollection, retrieves each object's binary data (ObjectData or FullObjectBin), determines a suitable file extension, and saves the content as uniquely named files in the target folder.
class ExtractOleObjects
{
    static void Main()
    {
        // Path to the source workbook
        string workbookPath = "input.xlsx";

        // Folder where extracted OLE objects will be saved
        string outputFolder = "ExtractedOleObjects";

        // Load the workbook (load rule)
        Workbook workbook = new Workbook(workbookPath);

        // Ensure the output directory exists
        Directory.CreateDirectory(outputFolder);

        // Iterate through each worksheet
        for (int sheetIdx = 0; sheetIdx < workbook.Worksheets.Count; sheetIdx++)
        {
            Worksheet sheet = workbook.Worksheets[sheetIdx];
            OleObjectCollection oleObjects = sheet.OleObjects;

            // Iterate through each OLE object in the worksheet
            for (int oleIdx = 0; oleIdx < oleObjects.Count; oleIdx++)
            {
                OleObject ole = oleObjects[oleIdx];

                // Try to get the embedded data; fallback to FullObjectBin if needed
                byte[] data = ole.ObjectData;
                if (data == null || data.Length == 0)
                {
                    data = ole.FullObjectBin;
                }

                // Skip if there is no data to extract
                if (data == null || data.Length == 0)
                    continue;

                // Determine file extension
                string ext = Path.GetExtension(ole.ObjectSourceFullName);
                if (string.IsNullOrEmpty(ext))
                {
                    // If source name is not available, default to .bin
                    ext = ".bin";
                }

                // Build the output file name
                string fileName = $"Sheet{sheetIdx}_Ole{oleIdx}{ext}";
                string filePath = Path.Combine(outputFolder, fileName);

                // Write the extracted data to disk
                File.WriteAllBytes(filePath, data);
            }
        }
    }
}
