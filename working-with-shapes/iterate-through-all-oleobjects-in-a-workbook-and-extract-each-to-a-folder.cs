// Title: Extract all embedded OLE objects from an Excel workbook to a folder with Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx file with Aspose.Cells, loops through each worksheet, and writes the ObjectData of every OleObject to a specified directory. | Modify the extraction logic to use the OleObject.Name as the filename and fall back to a sequential name when the Name property is empty. | Add error handling that logs OleObjects without data and continues processing the remaining objects.
// Common Searches: how to export embedded OLE objects from an Excel file using Aspose.Cells C# | c# iterate OleObject collection in workbook and save to disk | Aspose.Cells extract binary data of OLE objects to folder | save each OleObject from .xlsx as separate file with Aspose.Cells | handle missing OLE object data when extracting from Excel in .NET
// Tags: extract OLE objects Aspose.Cells | save OleObject binary data C# | iterate worksheet OleObject collection | fallback filename for unnamed OleObject | log missing OLE data Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// The program loads an Excel workbook using Aspose.Cells, iterates through every worksheet's OleObjects collection, and writes each object's binary data to a target folder, naming files from the OleObject.Name or a generated identifier and logging any objects without data.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the source workbook
            string workbookPath = "input.xlsx";

            // Verify that the workbook file exists
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"Error: Workbook file not found at '{workbookPath}'.");
                return;
            }

            // Destination folder for extracted OLE objects
            string outputFolder = "ExtractedOleObjects";

            // Ensure the output folder exists
            Directory.CreateDirectory(outputFolder);

            // Load the workbook
            Workbook workbook = new Workbook(workbookPath);

            // Counter for generating unique file names when needed
            int oleCounter = 0;

            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through each OleObject in the current worksheet
                foreach (OleObject ole in sheet.OleObjects)
                {
                    try
                    {
                        // Base name: use the OLE object's name if available
                        string baseName = !string.IsNullOrEmpty(ole.Name)
                            ? ole.Name
                            : $"OleObject_{oleCounter}";

                        // Default file extension (binary data)
                        string extension = ".bin";

                        // Combine folder, name, and extension to get the full output path
                        string outputPath = Path.Combine(outputFolder, baseName + extension);

                        // Save the OLE object data to the file system (if data is available)
                        if (ole.ObjectData != null && ole.ObjectData.Length > 0)
                        {
                            File.WriteAllBytes(outputPath, ole.ObjectData);
                            Console.WriteLine($"Extracted: {outputPath}");
                        }
                        else
                        {
                            Console.WriteLine($"Warning: OleObject '{baseName}' has no data to save.");
                        }

                        oleCounter++;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to extract an OLE object: {ex.Message}");
                    }
                }
            }

            Console.WriteLine("OLE object extraction completed.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
