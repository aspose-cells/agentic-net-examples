using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class ExtractOleObjects
{
    static void Main()
    {
        // Path to the source Excel file containing embedded OLE objects
        string sourceFile = "InputWorkbook.xlsx";

        // Folder where extracted OLE files will be saved
        string outputFolder = "ExtractedOleObjects";

        // Ensure the output directory exists
        Directory.CreateDirectory(outputFolder);

        // Load the workbook
        Workbook workbook = new Workbook(sourceFile);

        // Iterate through all worksheets
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Iterate through all OLE objects in the current worksheet
            for (int i = 0; i < sheet.OleObjects.Count; i++)
            {
                OleObject ole = sheet.OleObjects[i];

                // Get the full binary data of the embedded OLE object
                byte[] oleData = ole.FullObjectBin;

                if (oleData == null || oleData.Length == 0)
                    continue; // No data to extract

                // Determine a file name for the extracted object
                // Prefer the original source file name if available; otherwise generate a unique name
                string fileName = Path.GetFileName(ole.ObjectSourceFullName);
                if (string.IsNullOrEmpty(fileName))
                {
                    // Use the OLE object's index and appropriate extension based on its format type
                    string ext = ole.FileFormatType.ToString().ToLower(); // e.g., "docx", "xlsx"
                    fileName = $"OleObject_{sheet.Name}_{i}{(ext.StartsWith("unknown") ? "" : "." + ext)}";
                }

                // Combine output folder and file name
                string outputPath = Path.Combine(outputFolder, fileName);

                // Write the binary data to disk
                File.WriteAllBytes(outputPath, oleData);

                Console.WriteLine($"Extracted OLE object to: {outputPath}");
            }
        }
    }
}