// Title: How to update external named range paths in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that loads a workbook with Aspose.Cells, iterates over all defined names, replaces any old external file path in the RefersTo formula with a new path, and saves the workbook. | Create a reusable method UpdateExternalNamedRanges(string workbookPath, string oldPath, string newPath, string outputPath) that updates external references in named ranges, validates file existence, creates missing directories, and writes the updated file. | Provide a robust C# example that handles exceptions, verifies the source workbook, updates external links in both workbook‑ and worksheet‑scoped names, ensures the output folder exists, and saves the modified workbook.
// Common Searches: Aspose.Cells C# change external file reference in named range after moving source workbook | Update RefersTo formula path for external named ranges using .NET | Programmatically fix broken external links in Excel defined names with Aspose.Cells
// Tags: Aspose.Cells update external named range path | C# modify RefersTo external link | Aspose.Cells defined names external reference | Excel workbook external file relocation .NET | Aspose.Cells save workbook after link update

using System;
using System.IO;
using Aspose.Cells;

namespace ExternalReferenceUpdater
{
    // The sample loads an existing Excel workbook, scans every defined name (both workbook‑ and worksheet‑scoped), replaces occurrences of a specified old external file path in the RefersTo property with a new path, ensures the destination directory exists, and saves the updated workbook, all with comprehensive error handling.
    class Program
    {
        static void Main(string[] args)
        {
            // Paths for the workbook and external files
            string workbookPath = @"C:\Data\MyWorkbook.xlsx";
            string updatedWorkbookPath = @"C:\Data\MyWorkbook_Updated.xlsx";
            string oldExternalPath = @"C:\OldFolder\ExternalFile.xlsx";
            string newExternalPath = @"D:\NewFolder\ExternalFile.xlsx";

            try
            {
                // Verify that the source workbook exists
                if (!File.Exists(workbookPath))
                {
                    Console.WriteLine($"Workbook not found: {workbookPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook;
                try
                {
                    workbook = new Workbook(workbookPath);
                }
                catch (Exception loadEx)
                {
                    Console.WriteLine($"Failed to load workbook: {loadEx.Message}");
                    return;
                }

                // Get all defined names (both workbook‑ and worksheet‑scoped)
                NameCollection allNames = workbook.Worksheets.Names;
                for (int i = 0; i < allNames.Count; i++)
                {
                    Name name = allNames[i];
                    string refersTo = name.RefersTo;

                    // Update external references that contain the old path
                    if (!string.IsNullOrEmpty(refersTo) &&
                        refersTo.IndexOf(oldExternalPath, StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        string updatedRef = refersTo.Replace(oldExternalPath, newExternalPath, StringComparison.OrdinalIgnoreCase);
                        name.RefersTo = updatedRef;
                    }
                }

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(updatedWorkbookPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the updated workbook
                try
                {
                    workbook.Save(updatedWorkbookPath);
                    Console.WriteLine($"Workbook saved successfully to: {updatedWorkbookPath}");
                }
                catch (Exception saveEx)
                {
                    Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An unexpected error occurred: {ex.Message}");
            }
        }
    }
}
