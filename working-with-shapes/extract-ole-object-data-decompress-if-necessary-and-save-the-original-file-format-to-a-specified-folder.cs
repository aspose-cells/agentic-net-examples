// Title: Extract embedded OLE objects from an Excel workbook and save them in their original file formats with Aspose.Cells for .NET
// Common Searches: C# code to extract OLE objects from an .xlsx file using Aspose.Cells | how to save embedded OLE object data to its original extension with Aspose.Cells .NET | retrieve raw ObjectData from worksheet OleObjectCollection in Aspose.Cells | export all OLE objects from Excel workbook to a folder using C# Aspose.Cells | Aspose.Cells extract embedded Word or PDF OLE objects from Excel
// Tags: Aspose.Cells OLE extraction workflow | saving OleObject.ObjectData as original file | inferring file type from OleObject.Name | iterating OleObjectCollection across worksheets | writing binary OLE payload to filesystem C# | handling embedded OLE content in Excel

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample loads a workbook, walks through each worksheet's OleObjectCollection, reads each OleObject's ObjectData, derives a filename from the object's Name (preserving the original extension when present), and writes the byte array to a specified output directory, logging successes and handling missing data or errors.
class OleObjectExtractor
{
    static void Main()
    {
        // Path to the source Excel workbook
        string sourceFile = @"C:\Input\Workbook.xlsx";

        // Folder where extracted OLE objects will be saved
        string outputFolder = @"C:\Output\OleObjects";

        // Verify that the source file exists
        if (!File.Exists(sourceFile))
        {
            Console.WriteLine($"Source file not found: {sourceFile}");
            return;
        }

        // Ensure the output directory exists
        Directory.CreateDirectory(outputFolder);

        Workbook workbook;
        try
        {
            // Load the workbook
            workbook = new Workbook(sourceFile);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to load workbook: {ex.Message}");
            return;
        }

        int oleIndex = 0;

        // Iterate through all worksheets in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            OleObjectCollection oleObjects = sheet.OleObjects;

            // Process each OLE object
            foreach (OleObject ole in oleObjects)
            {
                oleIndex++;

                // Determine a file name for the extracted object.
                // Use the OLE object's name if it contains an extension; otherwise generate one.
                string baseName = $"OleObject_{oleIndex}";
                string extension = "bin";

                if (!string.IsNullOrEmpty(ole.Name))
                {
                    string name = ole.Name;
                    int dotPos = name.LastIndexOf('.');
                    if (dotPos > -1 && dotPos < name.Length - 1)
                    {
                        baseName = name.Substring(0, dotPos);
                        extension = name.Substring(dotPos + 1);
                    }
                    else
                    {
                        baseName = name;
                    }
                }

                string fileName = $"{baseName}.{extension}";
                string savePath = Path.Combine(outputFolder, fileName);

                try
                {
                    // Retrieve the raw OLE object data (use ObjectData property)
                    byte[] data = ole.ObjectData;

                    if (data != null && data.Length > 0)
                    {
                        // Write the data to disk
                        File.WriteAllBytes(savePath, data);
                        Console.WriteLine($"Saved OLE object #{oleIndex} to \"{savePath}\".");
                    }
                    else
                    {
                        Console.WriteLine($"OLE object #{oleIndex} contains no data.");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to save OLE object #{oleIndex}: {ex.Message}");
                }
            }
        }

        Console.WriteLine($"Extraction complete. {oleIndex} OLE object(s) processed.");
    }
}
