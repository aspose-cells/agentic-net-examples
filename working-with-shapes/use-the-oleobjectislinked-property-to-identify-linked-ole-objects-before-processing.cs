// Title: Identify linked OLE objects and extract embedded OLE data from Excel workbooks using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells check if an OLE object is linked before extracting data | How to extract embedded OLE objects from an Excel file with Aspose.Cells .NET | Skip linked OLE objects while iterating worksheets using Aspose.Cells | Save embedded OLE object files to disk with Aspose.Cells in C# | Get source path of linked OLE objects using Aspose.Cells IsLink property
// Tags: Aspose.Cells identify linked OLE objects | extract embedded OLE data Aspose.Cells | OleObject.IsLink C# | save embedded OLE files Aspose.Cells | Excel workbook OLE object processing .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace OleObjectLinkChecker
{
    // The example loads an Excel workbook with Aspose.Cells, walks through each worksheet's OleObjectCollection, uses the OleObject.IsLink property to separate linked from embedded OLE objects, logs linked object details, extracts and writes embedded object data to a user‑specified folder (inferring file extensions from the ProgID), and finally reports counts of total, linked, and embedded OLE objects.
    class Program
    {
        static void Main(string[] args)
        {
            // Validate arguments
            if (args.Length < 1)
            {
                Console.WriteLine("Error: Input workbook path is required.");
                Console.WriteLine("Usage: OleObjectLinkChecker.exe <inputWorkbookPath> [outputFolder]");
                return;
            }

            string inputPath = args[0];
            string outputFolder = args.Length >= 2 ? args[1] : Path.Combine(Path.GetDirectoryName(inputPath) ?? "", "OleObjectData");

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" does not exist.");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(inputPath);
                Console.WriteLine($"Loaded workbook: {Path.GetFileName(inputPath)}");
                Console.WriteLine($"Worksheets count: {workbook.Worksheets.Count}");

                // Ensure output folder exists
                Directory.CreateDirectory(outputFolder);
                Console.WriteLine($"OLE object data will be saved to: {outputFolder}");

                int totalOleObjects = 0;
                int linkedOleObjects = 0;
                int embeddedOleObjects = 0;

                // Iterate through each worksheet
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    Console.WriteLine($"\nProcessing worksheet: {sheet.Name}");

                    OleObjectCollection oleObjects = sheet.OleObjects;
                    if (oleObjects == null || oleObjects.Count == 0)
                    {
                        Console.WriteLine("  No OLE objects found on this sheet.");
                        continue;
                    }

                    Console.WriteLine($"  OLE objects count: {oleObjects.Count}");

                    // Iterate through each OleObject
                    for (int i = 0; i < oleObjects.Count; i++)
                    {
                        OleObject ole = oleObjects[i];
                        totalOleObjects++;

                        // Identify if the OLE object is linked
                        bool isLinked = ole.IsLink;
                        if (isLinked)
                        {
                            linkedOleObjects++;
                            Console.WriteLine($"    [Linked] OleObject #{i + 1} - Label: \"{ole.Label}\"");
                            Console.WriteLine($"      Linked source: {ole.ObjectSourceFullName}");
                            // Linked objects typically do not contain embedded data, so we skip extraction.
                        }
                        else
                        {
                            embeddedOleObjects++;
                            Console.WriteLine($"    [Embedded] OleObject #{i + 1} - Label: \"{ole.Label}\"");
                            Console.WriteLine($"      ProgID: {ole.ProgID}");
                            Console.WriteLine($"      Size: {ole.Width}x{ole.Height} pixels");

                            // Extract embedded data if available
                            byte[] objectData = ole.ObjectData;
                            if (objectData != null && objectData.Length > 0)
                            {
                                // Determine a safe file name
                                string safeLabel = string.IsNullOrWhiteSpace(ole.Label) ? $"OleObject_{i + 1}" : ole.Label;
                                foreach (char c in Path.GetInvalidFileNameChars())
                                    safeLabel = safeLabel.Replace(c, '_');

                                // Use the ProgID or file format to guess an extension
                                string extension = GetExtensionFromProgID(ole.ProgID) ?? ".bin";

                                string outputFilePath = Path.Combine(outputFolder, $"{sheet.Name}_{safeLabel}{extension}");
                                File.WriteAllBytes(outputFilePath, objectData);
                                Console.WriteLine($"      Embedded data saved to: {outputFilePath}");
                            }
                            else
                            {
                                Console.WriteLine("      No embedded data found for this OLE object.");
                            }
                        }
                    }
                }

                // Summary
                Console.WriteLine("\n=== Processing Summary ===");
                Console.WriteLine($"Total OLE objects examined: {totalOleObjects}");
                Console.WriteLine($"Linked OLE objects: {linkedOleObjects}");
                Console.WriteLine($"Embedded OLE objects: {embeddedOleObjects}");
                Console.WriteLine("Processing completed successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Fatal error: {ex.Message}");
                Console.WriteLine("Stack Trace:");
                Console.WriteLine(ex.StackTrace);
            }
        }

        /// <summary>
        /// Attempts to map a ProgID to a common file extension.
        /// This is a best‑effort helper; if the ProgID is unknown, null is returned.
        /// </summary>
        private static string GetExtensionFromProgID(string progId)
        {
            if (string.IsNullOrWhiteSpace(progId))
                return null;

            // Common ProgID mappings
            switch (progId.ToLowerInvariant())
            {
                case "excel.sheet.8":
                case "excel.sheet":
                    return ".xls";
                case "excel.sheet.12":
                    return ".xlsx";
                case "word.document.8":
                case "word.document":
                    return ".doc";
                case "word.document.12":
                    return ".docx";
                case "powerpoint.show.8":
                case "powerpoint.show":
                    return ".ppt";
                case "powerpoint.show.12":
                    return ".pptx";
                case "package":
                    return ".zip";
                default:
                    // Unknown ProgID – return null to let caller decide default extension
                    return null;
            }
        }
    }
}
