// Title: C# – Remove ‘Temp’ VBA Module from All Excel Workbooks in a Folder with Aspose.Cells
// Description: A C# console utility that scans a directory for Excel files (xls, xlsx, xlsm, xlsb), detects VBA projects, deletes any module named "Temp", and overwrites the original workbooks using Aspose.Cells for .NET.
// Keywords: Aspose.Cells VBA module removal | C# delete Temp macro | batch Excel VBA cleanup | overwrite Excel workbook after VBA edit | macro‑enabled workbook processing
// Common Searches: how to delete a specific VBA module from multiple Excel files using Aspose.Cells | batch remove Temp macro from .xlsm files c# | overwrite original Excel files after removing VBA modules .NET | check for macros before editing workbook with Aspose.Cells | remove unwanted VBA code from a folder of workbooks
// Developer Intent: Delete every VBA module named "Temp" from each workbook in a given folder and save the modified files in place.
// Use Cases: Clean temporary or debug macros before publishing workbooks to end users. | Automate compliance by stripping prohibited VBA code from archived spreadsheets. | Prepare a batch of macro‑enabled files for migration to a macro‑free environment.
// AI Prompts: Write C# code with Aspose.Cells that removes a list of VBA modules (e.g., Temp, Debug) from all Excel files in a directory and logs each change. | Enhance the RemoveTempVbaModules example with detailed error handling for read‑only files, permission issues, and corrupted workbooks. | Create a PowerShell wrapper that calls the C# utility to process folders supplied via command‑line arguments.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsExamples
{
    // A C# console utility that scans a directory for Excel files (xls, xlsx, xlsm, xlsb), detects VBA projects, deletes any module named "Temp", and overwrites the original workbooks using Aspose.Cells for .NET.
    public class RemoveTempVbaModules
    {
        // Removes VBA modules named "Temp" from each workbook in the specified folder
        // and saves the modified workbooks, overwriting the original files.
        public static void Run(string folderPath)
        {
            if (string.IsNullOrWhiteSpace(folderPath) || !Directory.Exists(folderPath))
            {
                Console.WriteLine("Invalid or non‑existent folder path.");
                return;
            }

            // Get all Excel files in the folder (including macro‑enabled formats)
            string[] workbookFiles = Directory.GetFiles(folderPath, "*.*", SearchOption.TopDirectoryOnly);
            foreach (string filePath in workbookFiles)
            {
                // Process only known Excel extensions
                string ext = Path.GetExtension(filePath).ToLowerInvariant();
                if (ext != ".xls" && ext != ".xlsx" && ext != ".xlsm" && ext != ".xlsb")
                {
                    continue;
                }

                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found: {Path.GetFileName(filePath)}");
                    continue;
                }

                try
                {
                    // Load the workbook
                    using (Workbook workbook = new Workbook(filePath))
                    {
                        // Proceed only if the workbook contains macros/VBA project
                        if (workbook.HasMacro)
                        {
                            VbaModuleCollection modules = workbook.VbaProject.Modules;

                            // Check if a module named "Temp" exists
                            bool tempExists = false;
                            foreach (VbaModule module in modules)
                            {
                                if (module.Name.Equals("Temp", StringComparison.OrdinalIgnoreCase))
                                {
                                    tempExists = true;
                                    break;
                                }
                            }

                            if (tempExists)
                            {
                                modules.Remove("Temp");
                                // Save the workbook, overwriting the original file
                                workbook.Save(filePath);
                                Console.WriteLine($"Removed 'Temp' module and saved: {Path.GetFileName(filePath)}");
                            }
                            else
                            {
                                Console.WriteLine($"No 'Temp' module found in: {Path.GetFileName(filePath)}");
                            }
                        }
                        else
                        {
                            Console.WriteLine($"Workbook does not contain macros: {Path.GetFileName(filePath)}");
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{Path.GetFileName(filePath)}': {ex.Message}");
                }
            }
        }
    }

    // Entry point for the console application
    public class Program
    {
        public static void Main(string[] args)
        {
            try
            {
                string folderPath = args.Length > 0 ? args[0] : Directory.GetCurrentDirectory();
                Console.WriteLine($"Processing folder: {folderPath}");
                RemoveTempVbaModules.Run(folderPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }
    }
}
