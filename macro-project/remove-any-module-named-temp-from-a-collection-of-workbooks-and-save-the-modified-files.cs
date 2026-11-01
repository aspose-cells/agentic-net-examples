// Title: Remove a VBA module named "Temp" from multiple Excel workbooks using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# remove VBA module named Temp from workbook | batch delete specific VBA module in Excel files using Aspose.Cells .NET | C# code to iterate VBA modules collection and remove by name with Aspose.Cells | preserve original Excel file format when saving after VBA changes Aspose.Cells | handle missing Excel files while removing VBA modules with Aspose.Cells
// Tags: remove VBA module Aspose.Cells | batch process Excel workbooks C# | preserve original file format Aspose.Cells | case-insensitive module name removal C# | VBA modules collection iteration Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The example loads each workbook from a list of file paths, checks for a VBA project, and removes any module whose name matches "Temp" (case‑insensitive). It determines the original file format from the file extension and saves the workbook back, preserving the format while handling missing files and runtime errors.
public class WorkbookProcessor
{
    /// <param name="workbookPaths">Full file paths of the workbooks to process.</param>
    public void RemoveTempModules(IEnumerable<string> workbookPaths)
    {
        foreach (var path in workbookPaths)
        {
            // Skip non‑existent files to avoid FileNotFoundException.
            if (!File.Exists(path))
            {
                Console.WriteLine($"File not found: {path}");
                continue;
            }

            try
            {
                // Load the workbook from file.
                var workbook = new Workbook(path);

                // Ensure the workbook contains a VBA project.
                if (workbook.VbaProject != null && workbook.VbaProject.Modules != null)
                {
                    var modules = workbook.VbaProject.Modules;

                    // Iterate backwards so removal does not affect the index of remaining items.
                    for (int i = modules.Count - 1; i >= 0; i--)
                    {
                        // Check if the module name matches "Temp" (case‑insensitive).
                        if (string.Equals(modules[i].Name, "Temp", StringComparison.OrdinalIgnoreCase))
                        {
                            // Remove the matching module.
                            modules.RemoveAt(i);
                        }
                    }
                }

                // Determine the original file format to preserve it on save.
                SaveFormat format = GetSaveFormatFromExtension(Path.GetExtension(path));

                // Save the modified workbook, overwriting the original file.
                workbook.Save(path, format);
                Console.WriteLine($"Processed and saved: {path}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing '{path}': {ex.Message}");
            }
        }
    }

    private SaveFormat GetSaveFormatFromExtension(string extension)
    {
        switch (extension?.ToLowerInvariant())
        {
            case ".xls":
                return SaveFormat.Excel97To2003;
            case ".xlsx":
                return SaveFormat.Xlsx;
            case ".xlsm":
                return SaveFormat.Xlsm;
            case ".csv":
                return SaveFormat.Csv; // Updated to non‑obsolete enum
            case ".ods":
                return SaveFormat.Ods; // Updated to non‑obsolete enum
            default:
                // Default to XLSX if the extension is unrecognized.
                return SaveFormat.Xlsx;
        }
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        // Collect workbook paths from command‑line arguments or define them manually.
        var workbookPaths = new List<string>();

        if (args.Length > 0)
        {
            workbookPaths.AddRange(args);
        }
        else
        {
            // Example placeholder – replace with actual file paths as needed.
            // workbookPaths.Add(@"C:\Path\To\YourWorkbook.xlsx");
        }

        var processor = new WorkbookProcessor();
        processor.RemoveTempModules(workbookPaths);
    }
}
