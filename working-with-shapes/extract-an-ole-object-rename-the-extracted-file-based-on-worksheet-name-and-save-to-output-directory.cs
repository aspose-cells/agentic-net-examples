// Title: Extract OLE objects from each worksheet, rename with sheet name, and save to a folder using Aspose.Cells for .NET
// AI Prompts: Write C# code that opens an Excel workbook with Aspose.Cells, loops through every worksheet, pulls out each embedded OLE object, and writes it to a target directory using a filename that merges the sheet name and object sequence number. | Adjust the extraction routine to discover the original file extension of each OLE object via reflection on OleObjectData and append that extension to the saved filename. | Add robust error handling to skip OLE objects without OleObjectData or a Save method and log warnings during the extraction process using Aspose.Cells.
// Common Searches: how to export OLE objects from an Excel file using Aspose.Cells C# | save extracted OLE object with worksheet name Aspose.Cells .NET | retrieve original file extension of OleObjectData via reflection Aspose.Cells | extract all OLE objects from every sheet in a workbook with Aspose.Cells | Aspose.Cells batch OLE object extraction to folder
// Tags: extract OLE objects Aspose.Cells .NET | save OLE data with worksheet-based filenames | OleObjectData reflection file extension | multiple OLE objects extraction per worksheet | handle missing OleObjectData Aspose.Cells

using System;
using System.IO;
using System.Reflection;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample loads an Excel workbook, creates an output folder, iterates through each worksheet, and uses reflection to access each OleObject's data. It determines the appropriate file extension, builds a filename that combines the worksheet name with the object's index, and saves the extracted OLE content to the folder while gracefully handling missing properties or methods.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the source workbook
            string inputPath = "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Directory where extracted OLE files will be saved
            string outputDirectory = "ExtractedOleObjects";

            // Ensure the output directory exists
            Directory.CreateDirectory(outputDirectory);

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                int oleIndex = 0; // Counter for OLE objects on the current sheet

                // Iterate through all OLE objects on the current worksheet
                foreach (OleObject oleObject in sheet.OleObjects)
                {
                    oleIndex++;

                    try
                    {
                        // Use reflection to obtain OleObjectData (covers versions where the property may be missing)
                        PropertyInfo propInfo = oleObject.GetType().GetProperty("OleObjectData", BindingFlags.Public | BindingFlags.Instance);
                        object oleDataObj = propInfo?.GetValue(oleObject);

                        if (oleDataObj == null)
                        {
                            Console.WriteLine($"Warning: OleObjectData not available for object {oleIndex} on sheet \"{sheet.Name}\".");
                            continue;
                        }

                        // Determine file extension; default to .bin if unavailable
                        string extension = ".bin";
                        PropertyInfo extProp = oleDataObj.GetType().GetProperty("FileExtension", BindingFlags.Public | BindingFlags.Instance);
                        if (extProp != null)
                        {
                            string extVal = extProp.GetValue(oleDataObj) as string;
                            if (!string.IsNullOrEmpty(extVal))
                            {
                                extension = extVal.StartsWith(".") ? extVal : "." + extVal;
                            }
                        }

                        // Build a unique file name using the worksheet name and the OLE object's index
                        string extractedFileName = $"{sheet.Name}_{oleIndex}{extension}";
                        string extractedFilePath = Path.Combine(outputDirectory, extractedFileName);

                        // Invoke the Save method to write the OLE object data to the file system
                        MethodInfo saveMethod = oleDataObj.GetType().GetMethod("Save", new[] { typeof(string) });
                        if (saveMethod != null)
                        {
                            saveMethod.Invoke(oleDataObj, new object[] { extractedFilePath });
                        }
                        else
                        {
                            Console.WriteLine($"Warning: Save method not found for OLE object {oleIndex} on sheet \"{sheet.Name}\".");
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to extract OLE object at index {oleIndex} on sheet \"{sheet.Name}\": {ex.Message}");
                    }
                }
            }

            Console.WriteLine("OLE objects extraction completed successfully.");
        }
        catch (Exception ex)
        {
            // Log or display any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
