// Title: Batch delete the "Temp" VBA module from Excel workbooks with Aspose.Cells for .NET
// Description: Iterates through all Excel files in a folder, loads each workbook with Aspose.Cells, checks for VBA macros, removes the module named Temp via VbaProject.Modules.Remove, and saves the workbook in its original format to an output directory.
// Keywords: Aspose.Cells C# remove VBA module | VbaProject Modules.Remove example | batch delete Temp macro | process macro-enabled workbooks .NET | save workbook after VBA removal | Excel macro cleanup Aspose | C# delete VBA module programmatically
// Common Searches: How to remove a specific VBA module from multiple Excel files using Aspose.Cells | C# code to batch delete the Temp module in macro-enabled workbooks | Aspose.Cells example for VbaProject module removal | Save Excel files after stripping VBA code with Aspose | Remove VBA modules programmatically in .NET
// Developer Intent: The developer wants to delete the VBA module named "Temp" from each workbook in a folder and save the cleaned files.
// Use Cases: Cleaning temporary macro modules before distributing workbooks to end users. | Automating compliance by stripping unwanted VBA code from a batch of financial reports. | Preparing workbooks for macro‑restricted environments by removing all "Temp" modules.
// AI Prompts: Generate C# code using Aspose.Cells to remove a VBA module named "Temp" from all Excel files in a directory and save them to another folder. | Provide best‑practice error handling for VbaModuleCollection.Remove when processing multiple workbooks with Aspose.Cells. | Explain how to preserve the original file format while saving workbooks after modifying their VBA projects with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

// Iterates through all Excel files in a folder, loads each workbook with Aspose.Cells, checks for VBA macros, removes the module named Temp via VbaProject.Modules.Remove, and saves the workbook in its original format to an output directory.
class RemoveTempModules
{
    static void Main()
    {
        // Input folder containing the workbooks
        string inputFolder = @"C:\InputWorkbooks";
        // Output folder for the modified workbooks
        string outputFolder = @"C:\OutputWorkbooks";

        // Ensure the output folder exists
        Directory.CreateDirectory(outputFolder);

        // Get all Excel files (including macro‑enabled formats) from the input folder
        string[] workbookFiles = Directory.GetFiles(inputFolder, "*.*", SearchOption.TopDirectoryOnly);

        foreach (string filePath in workbookFiles)
        {
            // Load the workbook (uses the Workbook(string) constructor)
            Workbook workbook = new Workbook(filePath);

            // If the workbook contains VBA macros, try to remove the module named "Temp"
            if (workbook.HasMacro)
            {
                try
                {
                    // Remove the module by name (uses VbaModuleCollection.Remove(string))
                    workbook.VbaProject.Modules.Remove("Temp");
                }
                catch (Exception ex)
                {
                    // Log any errors but continue processing other workbooks
                    Console.WriteLine($"Error removing 'Temp' module from '{Path.GetFileName(filePath)}': {ex.Message}");
                }
            }

            // Save the modified workbook to the output folder (preserves original format)
            string destPath = Path.Combine(outputFolder, Path.GetFileName(filePath));
            workbook.Save(destPath);

            // Release resources
            workbook.Dispose();
        }

        Console.WriteLine("All workbooks processed.");
    }
}
