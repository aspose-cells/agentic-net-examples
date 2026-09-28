// Title: Merge two Excel workbooks in C# with Aspose.Cells while preserving named ranges and handling duplicate sheet or range names
// AI Prompts: Write C# code that uses Aspose.Cells to load two .xlsx workbooks, detect and rename any conflicting worksheet names, copy all worksheets from the second workbook into the first, and save the merged file. | Create C# logic with Aspose.Cells that merges the named ranges of two workbooks, automatically generates unique names for duplicates, copies comments, and writes the combined workbook to disk.
// Common Searches: c# aspose.cells merge two workbooks keep named ranges | how to avoid worksheet name collisions when combining Excel files with Aspose.Cells | rename duplicate named ranges during workbook merge asp.net | copy worksheets from one workbook to another using Aspose.Cells AddCopy method | merge excel files programmatically and resolve named range conflicts c#
// Tags: Aspose.Cells workbook merge with unique worksheet names | named range conflict handling Aspose.Cells | AddCopy worksheet duplication Aspose.Cells | C# merge Excel files preserving named ranges | automatic renaming of duplicate sheet names Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Cells;

namespace WorkbookMergeExample
{
    // The example loads two Excel workbooks, ensures worksheet names from the second workbook are unique by appending numeric suffixes, copies those worksheets into the first workbook using AddCopy, merges all named ranges while generating distinct names for any duplicates, preserves range comments, and saves the combined workbook as a new .xlsx file.
    class MergeWorkbooks
    {
        static void Main()
        {
            try
            {
                const string file1 = "Workbook1.xlsx";
                const string file2 = "Workbook2.xlsx";
                const string outputFile = "MergedWorkbook.xlsx";

                // Verify source files exist
                if (!File.Exists(file1) || !File.Exists(file2))
                {
                    Console.WriteLine("One or both source workbook files are missing.");
                    return;
                }

                // Load the two source workbooks
                Workbook wb1 = new Workbook(file1);
                Workbook wb2 = new Workbook(file2);

                // ---------- Ensure unique worksheet names ----------
                foreach (Worksheet ws in wb2.Worksheets)
                {
                    string originalName = ws.Name;
                    string uniqueName = originalName;

                    // Check for conflict with existing sheet names in wb1
                    if (wb1.Worksheets.Any(s => s.Name.Equals(uniqueName, StringComparison.OrdinalIgnoreCase)))
                    {
                        int suffix = 1;
                        while (wb1.Worksheets.Any(s => s.Name.Equals($"{originalName}_{suffix}", StringComparison.OrdinalIgnoreCase)))
                        {
                            suffix++;
                        }
                        uniqueName = $"{originalName}_{suffix}";
                    }

                    ws.Name = uniqueName; // apply the unique name
                }

                // ---------- Copy worksheets from wb2 into wb1 ----------
                foreach (Worksheet ws in wb2.Worksheets)
                {
                    // AddCopy expects the worksheet name (string) to copy from the source workbook
                    wb1.Worksheets.AddCopy(ws.Name);
                }

                // ---------- Merge named ranges ----------
                // Collect existing names from wb1 for quick lookup
                var existingNames = new HashSet<string>(wb1.Worksheets.Names.Select(n => n.Text), StringComparer.OrdinalIgnoreCase);

                foreach (Name srcName in wb2.Worksheets.Names)
                {
                    string targetName = srcName.Text;

                    // Resolve naming conflict by generating a unique name
                    if (existingNames.Contains(targetName))
                    {
                        int suffix = 1;
                        string baseName = targetName;
                        while (existingNames.Contains($"{baseName}_{suffix}"))
                        {
                            suffix++;
                        }
                        targetName = $"{baseName}_{suffix}";
                    }

                    // Add the (potentially renamed) name to the destination workbook
                    int index = wb1.Worksheets.Names.Add(targetName);
                    Name destName = wb1.Worksheets.Names[index];

                    // Copy essential properties
                    destName.RefersTo = srcName.RefersTo;   // address the name refers to
                    destName.Comment = srcName.Comment;    // optional comment

                    // Register the new name to avoid future conflicts
                    existingNames.Add(targetName);
                }

                // Save the merged workbook
                wb1.Save(outputFile);
                Console.WriteLine($"Workbooks merged successfully into '{outputFile}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
