// Title: C# code to delete the 'Temp' VBA module from all .xlsx workbooks in a folder using Aspose.Cells
// Common Searches: C# Aspose.Cells remove VBA module named Temp from multiple Excel files | batch delete specific VBA module in .xlsx workbooks with Aspose.Cells .NET | how to programmatically strip Temp VBA module from Excel workbooks using Aspose.Cells | iterate over a folder of .xlsx files and remove a VBA module with Aspose.Cells C# | skip password‑protected Excel files while removing VBA modules using Aspose.Cells
// Tags: Aspose.Cells delete VBA module Temp | batch VBA project cleanup .xlsx | C# overwrite workbook after VBA modification | skip password protected Excel files Aspose.Cells | iterate Excel files folder Aspose.Cells

using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Cells;

// The sample enumerates every .xlsx file in a specified directory, loads each workbook with Aspose.Cells, checks for a VBA project, removes any module whose name matches "Temp", and saves the workbook back to the original file. It includes error handling for password‑protected files and other exceptions, ensuring batch processing of workbooks without manual intervention.
class Program
{
    static void Main()
    {
        // Path to the folder containing the workbooks
        string folderPath = @"C:\Workbooks";

        // Verify that the folder exists
        if (!Directory.Exists(folderPath))
        {
            Console.WriteLine($"Folder not found: {folderPath}");
            return;
        }

        // Retrieve all Excel files in the folder (adjust pattern if needed)
        string[] workbookFiles = Directory.GetFiles(folderPath, "*.xlsx", SearchOption.TopDirectoryOnly);

        foreach (string filePath in workbookFiles)
        {
            // Ensure the file still exists (safety check)
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found, skipping: {filePath}");
                continue;
            }

            try
            {
                // Load the workbook (no password supplied; if required an exception will be thrown)
                Workbook workbook = new Workbook(filePath);

                // Process VBA modules only if a VBA project exists
                if (workbook.VbaProject != null && workbook.VbaProject.Modules != null && workbook.VbaProject.Modules.Count > 0)
                {
                    // Collect indices of modules named "Temp"
                    List<int> indicesToRemove = new List<int>();

                    for (int i = 0; i < workbook.VbaProject.Modules.Count; i++)
                    {
                        var module = workbook.VbaProject.Modules[i];
                        if (string.Equals(module.Name, "Temp", StringComparison.OrdinalIgnoreCase))
                        {
                            indicesToRemove.Add(i);
                        }
                    }

                    // Remove modules in reverse order to avoid index shifting
                    for (int i = indicesToRemove.Count - 1; i >= 0; i--)
                    {
                        workbook.VbaProject.Modules.RemoveAt(indicesToRemove[i]);
                    }
                }

                // Save the modified workbook (overwrites the original file)
                workbook.Save(filePath);
                Console.WriteLine($"Processed: {filePath}");
            }
            catch (CellsException ex)
            {
                // Handle password‑protected files or other CellsException cases
                if (ex.Message != null && ex.Message.IndexOf("password", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    Console.WriteLine($"Password required, skipping file: {filePath}");
                }
                else
                {
                    Console.WriteLine($"CellsException processing file '{filePath}': {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                // Log any other errors and continue processing remaining files
                Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
            }
        }
    }
}
