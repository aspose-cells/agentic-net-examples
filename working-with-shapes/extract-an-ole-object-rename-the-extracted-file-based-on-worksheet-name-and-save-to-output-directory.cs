using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace OleObjectExtractor
{
    class Program
    {
        static void Main(string[] args)
        {
            // Input Excel file containing OLE objects
            string inputFile = "input.xlsx";

            // Directory where extracted files will be saved
            string outputDir = "ExtractedOleObjects";

            // Ensure the output directory exists
            Directory.CreateDirectory(outputDir);

            // Load the workbook (using the provided load rule)
            Workbook workbook = new Workbook(inputFile);

            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Get the worksheet name to use in the extracted file name
                string sheetName = sheet.Name;

                // Iterate through each OLE object in the current worksheet
                for (int i = 0; i < sheet.OleObjects.Count; i++)
                {
                    OleObject ole = sheet.OleObjects[i];

                    // Get the embedded OLE object data as a byte array
                    byte[] oleData = ole.ObjectData;

                    // Determine a file extension for the extracted file
                    // Prefer the original source file name if available; otherwise use the FileFormatType enum
                    string extension = ".bin"; // default fallback
                    if (!string.IsNullOrEmpty(ole.ObjectSourceFullName))
                    {
                        extension = Path.GetExtension(ole.ObjectSourceFullName);
                        if (string.IsNullOrEmpty(extension))
                            extension = ".bin";
                    }
                    else
                    {
                        // Use the FileFormatType enum name as extension (e.g., "Xlsx" -> ".xlsx")
                        string formatName = ole.FileFormatType.ToString().ToLower();
                        extension = "." + formatName;
                    }

                    // Build the output file name: <WorksheetName>_Ole_<index><extension>
                    string outputFileName = $"{sheetName}_Ole_{i + 1}{extension}";
                    string outputPath = Path.Combine(outputDir, outputFileName);

                    // Write the OLE object data to the file (using the provided save rule)
                    File.WriteAllBytes(outputPath, oleData);

                    Console.WriteLine($"Extracted OLE object from sheet '{sheetName}' to '{outputPath}'.");
                }
            }

            // Optionally, save the workbook if any modifications were made (not required for extraction)
            // workbook.Save("output.xlsx");
        }
    }
}