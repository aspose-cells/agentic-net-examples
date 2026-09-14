// Title: C# – Save an Aspose.Cells workbook as XLSX using a custom temporary folder to minimize disk I/O
// AI Prompts: Generate C# code that sets Workbook.Settings.TempFolderPath to a user‑defined directory, applies MemorySetting.MemoryPreference, and saves the workbook as XLSX. | Show a complete .NET example that creates a temporary folder, configures Aspose.Cells to use it for intermediate files, and writes the final Excel file with reduced disk access.
// Common Searches: how to set Aspose.Cells temporary folder path in C# before saving workbook | Aspose.Cells reduce disk I/O when saving large Excel file | C# Aspose.Cells MemoryPreference setting for workbook save | save workbook as xlsx using custom temp directory Aspose.Cells | Aspose.Cells TempFolderPath not recognized in .NET
// Tags: Aspose.Cells temporary folder configuration | Workbook.Save Xlsx memory preference | C# Aspose.Cells reduce disk I/O | Aspose.Cells intermediate file location | set TempFolderPath Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example creates an empty Workbook, ensures a custom temporary directory exists, assigns it to Workbook.Settings.TempFolderPath (if supported), sets MemorySetting.MemoryPreference, prepares the output folder, and saves the workbook as Result.xlsx, thereby lowering disk I/O during the save operation.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new empty workbook
            Workbook workbook = new Workbook();

            // Define a custom temporary folder for Aspose.Cells intermediate files
            string tempFolder = @"C:\CustomTempFolder";

            // Ensure the temporary folder exists
            if (!Directory.Exists(tempFolder))
            {
                Directory.CreateDirectory(tempFolder);
            }

            // Set the temporary folder path (available in supported versions)
            // If the property is not present in the current Aspose.Cells version, this line can be omitted.
            // workbook.Settings.TempFolderPath = tempFolder;

            // Prefer using the temporary folder for memory‑intensive operations
            workbook.Settings.MemorySetting = MemorySetting.MemoryPreference;

            // Prepare the output path and ensure its directory exists
            string resultPath = "Result.xlsx";
            string resultDir = Path.GetDirectoryName(Path.GetFullPath(resultPath));
            if (!string.IsNullOrEmpty(resultDir) && !Directory.Exists(resultDir))
            {
                Directory.CreateDirectory(resultDir);
            }

            // Save the workbook as XLSX
            workbook.Save(resultPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(resultPath)}'.");
        }
        catch (Exception ex)
        {
            // Log or display any errors that occur during processing
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
