// Title: C# – Extract Embedded OLE Objects from an Excel Workbook with Aspose.Cells
// Description: A console utility that loads an Excel file using Aspose.Cells, scans every worksheet for OleObjects, reads each object's FullObjectBin, determines a file name from ObjectSourceFullName (or maps FileFormatType to an extension), and writes the binary data to a user‑specified output folder.
// Keywords: Aspose.Cells | C# | .NET | OLE extraction | Excel embedded OLE | OleObject.FullObjectBin | ObjectSourceFullName | FileFormatType mapping | save OLE to disk | command‑line example | GitHub sample
// Common Searches: Aspose.Cells extract OLE objects from Excel | C# code to save embedded OLE files | Get original file name of OleObject Aspose | Map FileFormatType to file extension C# | Command line tool for OLE extraction in .NET
// Developer Intent: Programmatically pull every embedded OLE object out of an Excel workbook and store each as an individual file in a chosen directory.
// Use Cases: Bulk export of attached Word, PDF, or PowerPoint files from a template workbook for archival. | Automated retrieval of documents embedded in financial reports for downstream processing. | Generating a folder of source files from OLE objects to feed a document‑management system.
// AI Prompts: Generate a C# method that iterates through all worksheets in a Workbook, extracts each OleObject's binary data, and saves it to a target folder while preserving the original filename when available. | Create a C# helper that converts Aspose.Cells FileFormatType values to common file extensions for OLE extraction. | Explain how to fall back to OleObject.SourceFullName when ObjectSourceFullName is empty and ensure unique filenames for extracted objects.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace OleObjectExtractor
{
    // A console utility that loads an Excel file using Aspose.Cells, scans every worksheet for OleObjects, reads each object's FullObjectBin, determines a file name from ObjectSourceFullName (or maps FileFormatType to an extension), and writes the binary data to a user‑specified output folder.
    class Program
    {
        static void Main(string[] args)
        {
            // Validate arguments: first = input workbook path, second = output folder
            if (args.Length < 2)
            {
                Console.WriteLine("Usage: OleObjectExtractor <inputWorkbookPath> <outputFolder>");
                return;
            }

            string inputPath = args[0];
            string outputFolder = args[1];

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Ensure the output folder exists
            if (!Directory.Exists(outputFolder))
                Directory.CreateDirectory(outputFolder);

            try
            {
                // Load the workbook (lifecycle load rule)
                Workbook workbook = new Workbook(inputPath);

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
                        {
                            Console.WriteLine($"Worksheet '{sheet.Name}', OLE object #{i} has no data.");
                            continue;
                        }

                        // Determine a file name for the extracted object
                        string fileName;

                        // Prefer the original source file name if available
                        if (!string.IsNullOrEmpty(ole.ObjectSourceFullName))
                        {
                            fileName = Path.GetFileName(ole.ObjectSourceFullName);
                        }
                        else if (!string.IsNullOrEmpty(ole.SourceFullName))
                        {
                            // Fallback to the obsolete property (still usable)
                            fileName = Path.GetFileName(ole.SourceFullName);
                        }
                        else
                        {
                            // If no source name, create a generic name with appropriate extension based on FileFormatType
                            string extension = GetExtensionFromFormat(ole.FileFormatType);
                            fileName = $"OleObject_{sheet.Index}_{i}{extension}";
                        }

                        // Combine with output folder
                        string outputPath = Path.Combine(outputFolder, fileName);

                        // Write the binary data to disk
                        File.WriteAllBytes(outputPath, oleData);
                        Console.WriteLine($"Extracted OLE object to: {outputPath}");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Helper method to map FileFormatType to a common file extension
        private static string GetExtensionFromFormat(FileFormatType format)
        {
            switch (format)
            {
                case FileFormatType.Docx:
                case FileFormatType.Doc:
                    return ".docx";
                case FileFormatType.Xlsx:
                    return ".xlsx";
                case FileFormatType.Pptx:
                case FileFormatType.Ppt:
                    return ".pptx";
                case FileFormatType.Pdf:
                    return ".pdf";
                default:
                    return ".bin";
            }
        }
    }
}
