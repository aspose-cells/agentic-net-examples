// Title: C# code to extract embedded OLE objects from an Excel workbook and save them to a directory using Aspose.Cells
// Common Searches: how to extract embedded OLE objects from an .xlsx file with Aspose.Cells in C# | Aspose.Cells C# extract OLE object data and write to disk | save each OLE object from an Excel worksheet as a separate file using Aspose.Cells | iterate worksheets and export embedded OLE objects with Aspose.Cells for .NET | determine file extension of an OLE object when extracting via Aspose.Cells
// Tags: Aspose.Cells OLE object extraction | C# OleObjectCollection iteration | FileFormatType to file extension Aspose.Cells | write embedded OLE bytes to disk C# | export Excel OLE objects to folder

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace OleObjectExtractor
{
    // The sample loads an Excel workbook with Aspose.Cells, loops through every worksheet's OleObjectCollection, retrieves each object's binary data, creates a safe file name from its label (or a generated fallback), maps the object's FileFormatType to an appropriate file extension, and writes the data to a user‑specified output folder while handling missing data and I/O errors.
    class Program
    {
        /// <param name="args">
        /// args[0] - Path to the source Excel file (e.g., "C:\Temp\sample.xlsx").
        /// args[1] - Path to the output folder where extracted files will be written (e.g., "C:\Temp\ExtractedOle").
        /// </param>
        static void Main(string[] args)
        {
            // Validate arguments
            if (args.Length < 2)
            {
                Console.WriteLine("Usage: OleObjectExtractor <input_excel_path> <output_folder_path>");
                return;
            }

            string inputPath = args[0];
            string outputFolder = args[1];

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file not found: {inputPath}");
                return;
            }

            // Ensure output folder exists (create if necessary)
            try
            {
                Directory.CreateDirectory(outputFolder);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: Unable to create output folder '{outputFolder}'. {ex.Message}");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(inputPath);
                Console.WriteLine($"Workbook loaded: {inputPath}");

                // Iterate through all worksheets
                for (int wsIndex = 0; wsIndex < workbook.Worksheets.Count; wsIndex++)
                {
                    Worksheet sheet = workbook.Worksheets[wsIndex];
                    OleObjectCollection oleObjects = sheet.OleObjects;

                    if (oleObjects.Count == 0)
                        continue; // No OLE objects on this sheet

                    Console.WriteLine($"Worksheet \"{sheet.Name}\" contains {oleObjects.Count} OLE object(s).");

                    // Process each OLE object
                    for (int oleIndex = 0; oleIndex < oleObjects.Count; oleIndex++)
                    {
                        OleObject ole = oleObjects[oleIndex];

                        // Retrieve the embedded data
                        byte[] data = ole.ObjectData;
                        if (data == null || data.Length == 0)
                        {
                            Console.WriteLine($"  OLE object #{oleIndex + 1} has no data; skipping.");
                            continue;
                        }

                        // Determine a file name
                        string label = ole.Label;
                        if (string.IsNullOrWhiteSpace(label))
                        {
                            // Fallback to a generated name
                            label = $"OleObject_{wsIndex + 1}_{oleIndex + 1}";
                        }

                        // Clean label to be a valid file name
                        foreach (char invalid in Path.GetInvalidFileNameChars())
                        {
                            label = label.Replace(invalid, '_');
                        }

                        // Determine extension based on the stored file format type
                        string extension = GetExtensionFromFileFormat(ole.FileFormatType);
                        string outputFilePath = Path.Combine(outputFolder, $"{label}{extension}");

                        // Write the file safely
                        try
                        {
                            File.WriteAllBytes(outputFilePath, data);
                            Console.WriteLine($"  Extracted OLE object #{oleIndex + 1} to: {outputFilePath}");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"  Failed to write OLE object #{oleIndex + 1}: {ex.Message}");
                        }
                    }
                }

                Console.WriteLine("Extraction completed successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during extraction: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
            }
        }

        private static string GetExtensionFromFileFormat(FileFormatType format)
        {
            // Use the enum name as the basis for the extension.
            // Example: Xlsx -> .xlsx, Pdf -> .pdf, etc.
            string name = format.ToString().ToLowerInvariant();

            // Some enum names already include the dot-less extension.
            // Prepend a dot to form a proper file extension.
            return "." + name;
        }
    }
}
