// Title: C# – Extract Embedded OLE Objects from an Excel Workbook to a Folder Using Aspose.Cells
// Description: Loads an Excel file with Aspose.Cells, walks through each worksheet’s OleObjects collection, reads each object's FullObjectBin, maps its FileFormatType to a proper file extension, and writes the binary data to uniquely named files in a specified output directory. The workbook can then be saved unchanged.
// Keywords: Aspose.Cells OLE extraction | C# extract embedded OLE from Excel | OleObject.FullObjectBin | map FileFormatType to extension | save OLE objects to disk | Aspose.Cells .NET example | export embedded files from XLSX
// Common Searches: how to extract OLE objects from Excel with Aspose.Cells | save embedded Word/PDF files from workbook C# | Aspose.Cells get binary data of OleObject | C# export OLE objects to folder | determine file extension from Aspose.Cells FileFormatType
// Developer Intent: Retrieve every embedded OLE object in a workbook and write each one as a separate file to a chosen folder.
// Use Cases: Bulk export of embedded documents (Word, PDF, PowerPoint) for downstream processing. | Create a backup of all OLE content before performing mass edits on a template workbook. | Feed extracted files into OCR, conversion, or archival pipelines.
// AI Prompts: Generate C# code that extracts OLE objects from an Excel file with Aspose.Cells and logs unknown formats. | Propose improvements to the GetExtensionFromFileFormat method to cover additional Office and image types. | Explain how to read OLE object metadata (name, progID) alongside its binary content using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Loads an Excel file with Aspose.Cells, walks through each worksheet’s OleObjects collection, reads each object's FullObjectBin, maps its FileFormatType to a proper file extension, and writes the binary data to uniquely named files in a specified output directory. The workbook can then be saved unchanged.
class ExtractOleObjects
{
    static void Main()
    {
        // Path to the source workbook containing embedded OLE objects
        string inputWorkbookPath = "input.xlsx";

        // Folder where extracted OLE files will be saved
        string outputFolder = "ExtractedOleObjects";

        // Verify that the input workbook exists to avoid FileNotFoundException
        if (!File.Exists(inputWorkbookPath))
        {
            Console.WriteLine($"Input workbook not found: {inputWorkbookPath}");
            return;
        }

        Workbook workbook = null;
        try
        {
            // Load the workbook (load rule)
            workbook = new Workbook(inputWorkbookPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to load workbook: {ex.Message}");
            return;
        }

        // Ensure the output directory exists
        if (!Directory.Exists(outputFolder))
            Directory.CreateDirectory(outputFolder);

        int oleIndex = 0;

        // Iterate through all worksheets and their OLE objects
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            foreach (OleObject ole in sheet.OleObjects)
            {
                // Get the full binary data of the embedded OLE object
                byte[] oleData = ole.FullObjectBin;

                if (oleData != null && oleData.Length > 0)
                {
                    // Determine a suitable file extension based on the OLE object's format
                    string extension = GetExtensionFromFileFormat(ole.FileFormatType);

                    // Build a unique file name for each extracted object
                    string fileName = $"OleObject_{oleIndex}{extension}";
                    string fullPath = Path.Combine(outputFolder, fileName);

                    try
                    {
                        // Write the binary data to disk
                        File.WriteAllBytes(fullPath, oleData);
                        Console.WriteLine($"Extracted OLE object to: {fullPath}");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to write OLE object file: {ex.Message}");
                    }

                    oleIndex++;
                }
            }
        }

        // Optionally save the workbook unchanged (save rule)
        try
        {
            workbook.Save("output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to save workbook: {ex.Message}");
        }
    }

    // Helper method to map FileFormatType to a common file extension
    private static string GetExtensionFromFileFormat(FileFormatType format)
    {
        switch (format)
        {
            case FileFormatType.Doc:
            case FileFormatType.Docx:
                return ".docx";
            case FileFormatType.Xlsx:
                return ".xlsx";
            case FileFormatType.Xlsb:
            case FileFormatType.Xlsm:
                return ".xlsb";
            case FileFormatType.Ppt:
            case FileFormatType.Pptx:
                return ".pptx";
            case FileFormatType.Pdf:
                return ".pdf";
            case FileFormatType.Rtf:
                return ".rtf";
            case FileFormatType.Html:
                return ".html";
            default:
                // Fallback generic binary extension
                return ".bin";
        }
    }
}
