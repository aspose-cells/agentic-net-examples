// Title: How to extract embedded OLE objects from an Excel file and save them to a folder using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that opens a workbook with Aspose.Cells, enumerates all worksheets, extracts each OleObject's ObjectData, determines the correct file extension, and writes the bytes to a specified output directory. | Create a utility method in C# that maps Aspose.Cells OleObject.FileFormatType values to common file extensions for use during OLE extraction. | Add robust error handling to the extraction loop to skip objects with missing data, log the sheet name and object index, and continue processing remaining OLE objects.
// Common Searches: aspnet extract ole objects from xlsx using Aspose.Cells | c# save embedded OLE objects from Excel workbook to disk | determine file extension from OleObject.FileFormatType Aspose.Cells example | how to loop through worksheets and extract OLE objects with Aspose.Cells | Aspose.Cells OLE object extraction error handling
// Tags: extract OLE objects Aspose.Cells C# | OleObject binary extraction to file | map OleObject FileFormatType to extension | save embedded OLE data Aspose.Cells | worksheet iteration OLE extraction

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// The sample loads an Excel workbook with Aspose.Cells, iterates each worksheet, retrieves the binary data of every embedded OleObject, determines an appropriate file extension based on its FileFormatType, and writes the data to a designated output folder using uniquely constructed filenames, with error handling for missing data.
class Program
{
    static void Main()
    {
        // Path to the workbook containing embedded OLE objects
        string inputFile = @"C:\Input\sample.xlsx";

        // Folder where extracted OLE objects will be saved
        string outputFolder = @"C:\Output\OleObjects";

        // Ensure the output directory exists
        Directory.CreateDirectory(outputFolder);

        // Verify that the input file exists before attempting to load it
        if (!File.Exists(inputFile))
        {
            Console.WriteLine($"Input file not found: {inputFile}");
            return;
        }

        try
        {
            // Load the workbook (lifecycle rule: load)
            Workbook workbook = new Workbook(inputFile);

            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Access the collection of OLE objects on the current sheet
                OleObjectCollection oleObjects = sheet.OleObjects;

                // Process each OLE object
                for (int i = 0; i < oleObjects.Count; i++)
                {
                    OleObject ole = oleObjects[i];

                    try
                    {
                        // Retrieve the raw binary data of the OLE object
                        byte[] data = ole.ObjectData;

                        if (data == null || data.Length == 0)
                        {
                            Console.WriteLine($"No data found for OLE object {i} on sheet {sheet.Name}.");
                            continue;
                        }

                        // Determine a suitable file extension based on the OLE object's format
                        string extension = GetExtensionFromOleObject(ole);

                        // Build a unique file name for the extracted object
                        string fileName = $"{Path.GetFileNameWithoutExtension(inputFile)}_{sheet.Name}_Ole_{i}{extension}";
                        string outputPath = Path.Combine(outputFolder, fileName);

                        // Write the binary data to disk (lifecycle rule: save)
                        File.WriteAllBytes(outputPath, data);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to extract OLE object {i} on sheet {sheet.Name}: {ex.Message}");
                    }
                }
            }

            Console.WriteLine("OLE objects extraction completed successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred while processing the workbook: {ex.Message}");
        }
    }

    // Maps the OLE object's file format to a common file extension
    static string GetExtensionFromOleObject(OleObject ole)
    {
        // Use the string representation of the format to avoid dependency on specific enum members
        string fmt = ole.FileFormatType.ToString().ToLowerInvariant();

        if (fmt.Contains("excel"))
            return ".xls";
        if (fmt.Contains("word"))
            return ".doc";
        if (fmt.Contains("powerpoint"))
            return ".ppt";
        if (fmt.Contains("pdf"))
            return ".pdf";
        if (fmt.Contains("jpeg") || fmt.Contains("jpg"))
            return ".jpg";
        if (fmt.Contains("png"))
            return ".png";
        if (fmt.Contains("bmp"))
            return ".bmp";

        // Default binary extension
        return ".bin";
    }
}
