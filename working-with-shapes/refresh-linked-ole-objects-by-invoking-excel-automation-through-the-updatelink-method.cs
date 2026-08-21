// Title: C# – Refresh Linked OLE Objects in Excel using Aspose.Cells UpdateLinkedDataSource
// Description: This Aspose.Cells C# example loads an Excel workbook, finds every linked OLE object, enables its AutoUpdate flag, gathers the source file paths, opens each external workbook, refreshes the links with Workbook.UpdateLinkedDataSource, and saves the updated file.
// Keywords: Aspose.Cells | C# | Refresh OLE links | UpdateLinkedDataSource | OleObject.AutoUpdate | linked OLE objects | Excel automation | external workbook | GitHub sample | code example | Excel OLE refresh
// Common Searches: how to refresh linked OLE objects with Aspose.Cells | Aspose.Cells update OLE links C# | C# code to refresh OLE links in Excel | UpdateLinkedDataSource example Aspose | set OleObject AutoUpdate true Aspose.Cells | refresh OLE objects after source change
// Developer Intent: Programmatically refresh all linked OLE objects in an Excel workbook so they display the latest data from their source files.
// Use Cases: Ensure OLE charts and embedded documents reflect recent changes before sharing the workbook. | Automate batch updating of OLE links across multiple workbooks. | Validate and repair broken OLE source paths during a data pipeline. | Add a CI step that refreshes OLE links in generated reports.
// AI Prompts: Generate C# code using Aspose.Cells to locate linked OleObject items, enable AutoUpdate, and call UpdateLinkedDataSource. | Write error‑handling logic for missing OLE source files when refreshing links. | Provide a reusable method RefreshOleLinks(string inputPath, string outputPath) that returns a success status. | Create a PowerShell script that invokes the compiled .NET program to refresh OLE links in a folder of Excel files.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace RefreshOleLinksDemo
{
    // This Aspose.Cells C# example loads an Excel workbook, finds every linked OLE object, enables its AutoUpdate flag, gathers the source file paths, opens each external workbook, refreshes the links with Workbook.UpdateLinkedDataSource, and saves the updated file.
    class Program
    {
        static void Main()
        {
            try
            {
                const string inputPath = "InputWithOleLinks.xlsx";
                const string outputPath = "OutputWithRefreshedOleLinks.xlsx";

                // Verify that the input workbook exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Error: Input file not found – {inputPath}");
                    return;
                }

                // Load the workbook that contains linked OLE objects
                Workbook workbook = new Workbook(inputPath);

                // Collect unique source file names of linked OLE objects
                HashSet<string> sourceFiles = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    foreach (OleObject ole in sheet.OleObjects)
                    {
                        // Process only linked OLE objects
                        if (ole.IsLink)
                        {
                            // Ensure the OLE object is set to auto‑update
                            ole.AutoUpdate = true;

                            // Store the source file path for later loading
                            if (!string.IsNullOrEmpty(ole.ObjectSourceFullName))
                            {
                                sourceFiles.Add(ole.ObjectSourceFullName);
                            }
                        }
                    }
                }

                // Load each external workbook that is linked via OLE objects
                List<Workbook> externalWorkbooks = new List<Workbook>();
                foreach (string path in sourceFiles)
                {
                    if (File.Exists(path))
                    {
                        externalWorkbooks.Add(new Workbook(path));
                    }
                    else
                    {
                        Console.WriteLine($"Warning: Linked source file not found – {path}");
                    }
                }

                // Refresh the linked data sources (including OLE links) using Aspose.Cells API
                if (externalWorkbooks.Count > 0)
                {
                    workbook.UpdateLinkedDataSource(externalWorkbooks.ToArray());
                }

                // Save the updated workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
