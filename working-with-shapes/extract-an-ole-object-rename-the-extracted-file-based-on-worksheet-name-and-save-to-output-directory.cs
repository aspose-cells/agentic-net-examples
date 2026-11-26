// Title: Extract OLE objects from Excel worksheets, rename files with worksheet name, and save using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells extract embedded OLE objects from each worksheet | how to save extracted OLE data with original file extension in Aspose.Cells | rename extracted OLE files using worksheet name in a .NET application | Aspose.Cells determine OLE object file type before saving to disk | batch extract OLE objects from workbook and write to output folder C#
// Tags: extract OLE objects Aspose.Cells | rename extracted files with worksheet name | determine OLE file extension Aspose.Cells | save embedded OLE data C# | batch OLE extraction workbook Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace OleObjectExtractor
{
    // The example loads an Excel workbook with Aspose.Cells, loops through every worksheet, extracts each embedded OLE object's binary data, resolves an appropriate file extension using a format map, ProgID, or label fallback, builds a safe filename that combines the worksheet name and object index, and writes the file to a user‑specified output directory.
    class Program
    {
        // Map Aspose.Cells.FileFormatType values to common file extensions.
        private static readonly Dictionary<FileFormatType, string> FormatExtensionMap = new Dictionary<FileFormatType, string>
        {
            { FileFormatType.Xlsx, ".xlsx" },
            { FileFormatType.Xlsb, ".xlsb" },
            { FileFormatType.Xlsm, ".xlsm" },
            { FileFormatType.Csv, ".csv" },
            { FileFormatType.Pdf, ".pdf" },
            { FileFormatType.Doc, ".doc" },
            { FileFormatType.Docx, ".docx" },
            { FileFormatType.Ppt, ".ppt" },
            { FileFormatType.Pptx, ".pptx" },
            { FileFormatType.Png, ".png" },
            { FileFormatType.Gif, ".gif" },
            { FileFormatType.Bmp, ".bmp" },
            { FileFormatType.Tiff, ".tiff" }
            // Add more mappings as needed.
        };

        static void Main(string[] args)
        {
            // Expect two arguments: input workbook path and output directory.
            if (args.Length < 2)
            {
                Console.WriteLine("Usage: OleObjectExtractor <inputWorkbookPath> <outputDirectory>");
                return;
            }

            string inputPath = args[0];
            string outputDir = args[1];

            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" does not exist.");
                return;
            }

            try
            {
                // Ensure output directory exists.
                Directory.CreateDirectory(outputDir);

                // Load the workbook.
                Workbook workbook = new Workbook(inputPath);
                Console.WriteLine($"Workbook \"{Path.GetFileName(inputPath)}\" loaded successfully.");

                // Iterate through all worksheets.
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    string sheetName = sheet.Name;
                    OleObjectCollection oleObjects = sheet.OleObjects;

                    if (oleObjects == null || oleObjects.Count == 0)
                    {
                        Console.WriteLine($"Worksheet \"{sheetName}\" contains no OLE objects.");
                        continue;
                    }

                    Console.WriteLine($"Worksheet \"{sheetName}\" has {oleObjects.Count} OLE object(s).");

                    for (int i = 0; i < oleObjects.Count; i++)
                    {
                        OleObject ole = oleObjects[i];
                        try
                        {
                            // Retrieve the embedded data.
                            byte[] data = ole.ObjectData;
                            if (data == null || data.Length == 0)
                            {
                                Console.WriteLine($"  OLE object #{i + 1} has no data; skipping.");
                                continue;
                            }

                            // Determine file extension.
                            string extension = GetExtensionFromOleObject(ole);
                            if (string.IsNullOrEmpty(extension))
                            {
                                // Try to infer from label (e.g., "document.pdf").
                                string label = ole.Label;
                                if (!string.IsNullOrEmpty(label))
                                {
                                    extension = Path.GetExtension(label);
                                }
                                if (string.IsNullOrEmpty(extension))
                                {
                                    extension = ".bin";
                                }
                            }

                            // Build output file name: <WorksheetName>_Ole_<index><extension>
                            string safeSheetName = MakeFileSystemSafe(sheetName);
                            string fileName = $"{safeSheetName}_Ole_{i + 1}{extension}";
                            string outputPath = Path.Combine(outputDir, fileName);

                            // Write the file.
                            File.WriteAllBytes(outputPath, data);
                            Console.WriteLine($"  Extracted OLE object #{i + 1} to \"{outputPath}\".");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"  Error extracting OLE object #{i + 1}: {ex.Message}");
                        }
                    }
                }

                Console.WriteLine("Extraction completed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
            }
        }

        private static string GetExtensionFromOleObject(OleObject ole)
        {
            if (ole == null) return null;

            FileFormatType format = ole.FileFormatType;
            if (FormatExtensionMap.TryGetValue(format, out string ext))
            {
                return ext;
            }

            // Fallback: attempt to infer from ProgID.
            string progId = ole.ProgID?.ToLowerInvariant() ?? string.Empty;
            if (progId.Contains("excel"))
                return ".xlsx";
            if (progId.Contains("word"))
                return ".docx";
            if (progId.Contains("powerpoint"))
                return ".pptx";

            return null;
        }

        private static string MakeFileSystemSafe(string name)
        {
            foreach (char c in Path.GetInvalidFileNameChars())
            {
                name = name.Replace(c, '_');
            }
            return name;
        }
    }
}
