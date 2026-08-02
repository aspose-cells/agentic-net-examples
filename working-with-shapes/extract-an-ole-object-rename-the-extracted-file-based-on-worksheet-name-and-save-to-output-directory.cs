// Title: Extract and Rename Embedded OLE Objects from Excel Worksheets with Aspose.Cells (C#)
// Description: Load an Excel workbook, iterate through each worksheet, extract the binary data of every embedded OLE object, infer its file extension, create a filename that combines the worksheet name and object index, and write the file to a designated output folder. The workbook can be saved afterwards if needed.
// Keywords: Aspose.Cells OLE extraction | C# extract embedded OLE | rename OLE files by sheet name | determine OLE file extension | save OLE objects to folder | Excel embedded objects extraction | Aspose.Cells example C#
// Common Searches: how to extract OLE objects from Excel using Aspose.Cells | C# save embedded OLE files with worksheet name | Aspose.Cells get file extension of OleObject | extract and rename OLE objects in .NET | Aspose.Cells extract embedded PDFs from workbook
// Developer Intent: Programmatically pull each OLE object from an Excel file, name the extracted file with its worksheet identifier, and store it in a chosen directory.
// Use Cases: Archive all embedded documents from a multi‑sheet financial model, using sheet names for quick reference. | Migrate legacy reports by pulling PDFs, Word files, and images out of Excel before converting the workbook to a new format. | Create a backup of every OLE object in a template workbook prior to automated processing, ensuring original files remain untouched.
// AI Prompts: Write C# code that uses Aspose.Cells to loop through worksheets, extract each OleObject's data, infer its extension, and save the file with a name that includes the sheet name and object index. | Generate a helper method that maps OleObject.FileFormatType values to common file extensions, with a default .bin fallback. | Provide error‑handling snippets for missing input files, empty OLE data, and unsupported formats when extracting OLE objects from an Excel workbook.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace OleObjectExtractionDemo
{
    // Load an Excel workbook, iterate through each worksheet, extract the binary data of every embedded OLE object, infer its file extension, create a filename that combines the worksheet name and object index, and write the file to a designated output folder. The workbook can be saved afterwards if needed.
    class Program
    {
        static void Main()
        {
            // Input Excel file containing OLE objects
            string inputFile = "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Error: Input file '{inputFile}' was not found.");
                return;
            }

            // Directory where extracted OLE files will be saved
            string outputDirectory = "ExtractedOleObjects";

            // Ensure the output directory exists
            Directory.CreateDirectory(outputDirectory);

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(inputFile);

                // Iterate through each worksheet
                foreach (Worksheet worksheet in workbook.Worksheets)
                {
                    // Process each OLE object in the current worksheet
                    for (int i = 0; i < worksheet.OleObjects.Count; i++)
                    {
                        OleObject ole = worksheet.OleObjects[i];
                        byte[] oleData = ole.ObjectData;

                        // Skip if there is no embedded data
                        if (oleData == null || oleData.Length == 0)
                            continue;

                        // Determine file extension
                        string extension = GetExtensionFromOle(ole);

                        // Build a unique file name using worksheet name and object index
                        string fileName = $"{worksheet.Name}_OleObject_{i}{extension}";
                        string outputPath = Path.Combine(outputDirectory, fileName);

                        // Write the OLE data to the file system
                        File.WriteAllBytes(outputPath, oleData);
                    }
                }

                // Save the workbook (optional – can be omitted if not needed)
                workbook.Save("output.xlsx", SaveFormat.Xlsx);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Helper method to infer file extension for an OLE object
        private static string GetExtensionFromOle(OleObject ole)
        {
            // Try to use the original source file name if available
            string sourceFullName = ole.ObjectSourceFullName;
            if (!string.IsNullOrEmpty(sourceFullName))
            {
                string ext = Path.GetExtension(sourceFullName);
                if (!string.IsNullOrEmpty(ext))
                    return ext;
            }

            // Fallback: derive extension from the FileFormatType enum name
            string formatName = ole.FileFormatType.ToString().ToLowerInvariant();

            if (formatName.Contains("docx"))
                return ".docx";
            if (formatName.Contains("doc"))
                return ".doc";
            if (formatName.Contains("xlsx"))
                return ".xlsx";
            if (formatName.Contains("xls"))
                return ".xls";
            if (formatName.Contains("pptx"))
                return ".pptx";
            if (formatName.Contains("ppt"))
                return ".ppt";
            if (formatName.Contains("pdf"))
                return ".pdf";

            // Default binary extension
            return ".bin";
        }
    }
}
