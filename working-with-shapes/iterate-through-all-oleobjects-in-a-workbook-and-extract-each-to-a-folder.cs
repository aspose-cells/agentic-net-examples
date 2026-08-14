// Title: C# – Extract All OLE Objects from an Excel Workbook with Aspose.Cells
// Description: Loads an Excel workbook, iterates every worksheet and its OleObjects collection, determines a file name for linked or embedded objects, retrieves the binary via ObjectData or FullObjectBin, and saves each OLE object to a specified output folder.
// Keywords: Aspose.Cells | C# OLE extraction | extract OLE objects Excel | save embedded OLE files | linked OLE extraction .NET | OleObject iteration | Workbook OleObjects | export OLE binaries | Aspose.Cells example | GitHub C# Aspose.Cells
// Common Searches: How to extract OLE objects from Excel using Aspose.Cells C# | C# code to save embedded OLE files from a workbook | Export linked OLE objects to folder Aspose.Cells | Iterate OleObjects collection in .NET | Extract all OLE objects from all worksheets Aspose.Cells
// Developer Intent: Retrieve every OLE object (embedded or linked) from a workbook and write it to a chosen directory.
// Use Cases: Create an archive of all embedded charts, Word documents, and PDFs in a financial model for offline analysis. | Relocate linked OLE files before sharing a workbook to avoid broken references. | Generate a compliance‑ready backup of every OLE binary for audit trails. | Automate migration of legacy Excel reports that contain embedded objects to a document management system.
// AI Prompts: Write C# code that extracts both embedded and linked OLE objects from an Excel file using Aspose.Cells, with robust error handling. | Add logging to the OLE extraction routine that records file names, worksheet names, and missing‑data warnings. | Modify the example to group extracted objects by file extension into separate subfolders. | Provide a PowerShell script that calls the compiled C# extractor for batch processing of multiple workbooks.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Loads an Excel workbook, iterates every worksheet and its OleObjects collection, determines a file name for linked or embedded objects, retrieves the binary via ObjectData or FullObjectBin, and saves each OLE object to a specified output folder.
class ExtractOleObjects
{
    static void Main()
    {
        // Path to the source workbook
        string sourcePath = @"C:\Temp\InputWorkbook.xlsx";

        // Folder where extracted OLE files will be saved
        string outputFolder = @"C:\Temp\ExtractedOleObjects";

        try
        {
            // Verify that the source workbook exists
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source workbook not found: {sourcePath}");
                return;
            }

            // Ensure the output folder exists
            Directory.CreateDirectory(outputFolder);

            // Load the workbook
            Workbook workbook = new Workbook(sourcePath);

            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through each OLE object in the current worksheet
                foreach (OleObject ole in sheet.OleObjects)
                {
                    // Determine a file name for the extracted object
                    string fileName;

                    // If the OLE object is linked, use its source file name
                    if (ole.IsLink && !string.IsNullOrEmpty(ole.ObjectSourceFullName))
                    {
                        fileName = Path.GetFileName(ole.ObjectSourceFullName);
                    }
                    else
                    {
                        // For embedded objects, try to use the original source name if available
                        fileName = !string.IsNullOrEmpty(ole.ObjectSourceFullName)
                            ? Path.GetFileName(ole.ObjectSourceFullName)
                            : $"EmbeddedOle_{Guid.NewGuid()}.bin";
                    }

                    // Build the full path for the extracted file
                    string outputPath = Path.Combine(outputFolder, fileName);

                    // Extract the OLE data
                    byte[] data = ole.ObjectData;

                    // If ObjectData is null or empty, fall back to FullObjectBin
                    if (data == null || data.Length == 0)
                    {
                        data = ole.FullObjectBin;
                    }

                    // Write the data to the file system if we have any bytes
                    if (data != null && data.Length > 0)
                    {
                        File.WriteAllBytes(outputPath, data);
                        Console.WriteLine($"Extracted OLE object to: {outputPath}");
                    }
                    else
                    {
                        Console.WriteLine($"No data found for OLE object in sheet '{sheet.Name}'.");
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
