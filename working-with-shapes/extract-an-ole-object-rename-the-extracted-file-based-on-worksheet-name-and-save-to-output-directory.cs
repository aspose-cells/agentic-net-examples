// Title: C# – Extract and Save Embedded OLE Objects from Excel Worksheets with Aspose.Cells
// Description: Load an Excel workbook, loop through each worksheet, retrieve every embedded OLE object, infer the proper file extension from its FileFormatType, generate a filesystem‑safe name that includes the worksheet title and object index, and write the binary data to a designated output folder.
// Keywords: Aspose.Cells extract OLE | C# OLE object extraction | save embedded Excel objects | worksheet name file naming | FileFormatType to extension mapping | create output directory C# | batch export OLE from Excel | Aspose.Cells OLEObjectCollection
// Common Searches: how to extract OLE objects from Excel using Aspose.Cells C# | save embedded Word/PDF from Excel workbook | rename extracted OLE files with worksheet name | map Aspose.Cells FileFormatType to file extension | C# create safe file names from Excel sheet names
// Developer Intent: Programmatically pull every embedded OLE object from all worksheets, name each file with the sheet title and a sequential index, and store them in a chosen folder.
// Use Cases: Archive all embedded documents (Word, PowerPoint, PDF) from a multi‑sheet workbook for compliance audits. | Migrate legacy Excel files by separating embedded assets into individual files before conversion. | Automate content extraction for downstream processing, such as indexing extracted files in a search engine.
// AI Prompts: Generate C# code that uses Aspose.Cells to extract OLE objects, determine the correct extension from FileFormatType, and save each file with a name containing the worksheet title and object number. | Explain how to sanitize worksheet names for use in Windows file paths when exporting embedded OLE data with Aspose.Cells. | Extend the GetExtensionFromFormat method to include image (png, jpg) and audio (mp3, wav) formats embedded as OLE objects.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Load an Excel workbook, loop through each worksheet, retrieve every embedded OLE object, infer the proper file extension from its FileFormatType, generate a filesystem‑safe name that includes the worksheet title and object index, and write the binary data to a designated output folder.
class ExtractOleObjects
{
    static void Main()
    {
        // Path to the source Excel file
        string inputFile = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputFile))
        {
            Console.WriteLine($"Error: Input file \"{inputFile}\" not found.");
            return;
        }

        // Directory where extracted OLE files will be saved
        string outputDir = "ExtractedOleObjects";

        // Ensure the output directory exists
        Directory.CreateDirectory(outputDir);

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputFile);

            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Get the collection of OLE objects on the current worksheet
                OleObjectCollection oleObjects = sheet.OleObjects;

                // Process each OLE object
                for (int i = 0; i < oleObjects.Count; i++)
                {
                    OleObject ole = oleObjects[i];

                    // Retrieve the embedded OLE data as a byte array
                    byte[] data = ole.ObjectData;

                    // Determine a suitable file extension based on the object's format
                    string extension = GetExtensionFromFormat(ole.FileFormatType);

                    // Build a safe file name using the worksheet name and OLE index
                    string safeSheetName = MakeFileSystemSafe(sheet.Name);
                    string fileName = $"{safeSheetName}_Ole{i + 1}{extension}";
                    string outputPath = Path.Combine(outputDir, fileName);

                    // Write the OLE data to the file system
                    File.WriteAllBytes(outputPath, data);
                }
            }

            Console.WriteLine("OLE objects extraction completed successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Maps Aspose.Cells.FileFormatType to common file extensions
    static string GetExtensionFromFormat(FileFormatType format)
    {
        switch (format)
        {
            case FileFormatType.Xlsx: return ".xlsx";
            case FileFormatType.Xlsb: return ".xlsb";
            case FileFormatType.Docx: return ".docx";
            case FileFormatType.Doc: return ".doc";
            case FileFormatType.Pptx: return ".pptx";
            case FileFormatType.Ppt: return ".ppt";
            case FileFormatType.Pdf: return ".pdf";
            // Add more mappings as needed for supported formats
            default: return ".bin";
        }
    }

    // Replaces characters that are invalid in file names with an underscore
    static string MakeFileSystemSafe(string name)
    {
        foreach (char c in Path.GetInvalidFileNameChars())
        {
            name = name.Replace(c, '_');
        }
        return name;
    }
}
