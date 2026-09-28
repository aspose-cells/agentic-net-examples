// Title: Log timestamped additions, modifications, and deletions of Excel named ranges using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that captures the original RefersTo values of all named ranges in an Aspose.Cells workbook, applies changes, and appends a log entry with the current timestamp for each added, modified, or removed range to a text file. | Generate a C# example that compares a workbook's NameCollection before and after updates, detects differences, and creates audit entries like 'Added', 'Modified', 'Deleted' with old and new references and timestamps using Aspose.Cells.
// Common Searches: C# Aspose.Cells how to audit named range additions and deletions | track Excel named range reference changes with timestamp in .NET | detect modified named ranges in a workbook using Aspose.Cells C# | write named range change log to a text file with Aspose.Cells | compare original and current RefersTo values for Excel named ranges in C#
// Tags: Aspose.Cells named range audit logging | C# timestamped Excel named range change detection | compare NameCollection before after Aspose.Cells | log RefersTo updates to text file C# | detect added deleted named ranges Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace NamedRangeLogger
{
    // The example loads an Excel workbook with Aspose.Cells, records each named range's original RefersTo string, performs sample modifications (update, add, remove), then iterates the current NameCollection to identify added, modified, and deleted ranges. For every change it builds a timestamped log entry showing the old and new references and writes the entries to a text file.
    class Program
    {
        static void Main(string[] args)
        {
            // Paths for input workbook and log file
            string workbookPath = "input.xlsx";
            string logFilePath = "NamedRangeChanges.log";

            // Ensure the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"Error: Workbook file '{workbookPath}' not found.");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(workbookPath);
                NameCollection names = workbook.Worksheets.Names;

                // Capture original references of all named ranges
                var originalReferences = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                foreach (Name name in names)
                {
                    originalReferences[name.Text] = name.RefersTo;
                }

                // -----------------------------------------------------------------
                // Example modifications to named ranges (replace with actual logic)
                // -----------------------------------------------------------------

                // 1. Change the reference of an existing named range
                Name myRange = names["MyRange"];
                if (myRange != null)
                {
                    myRange.RefersTo = "=Sheet1!$C$1:$D$5";
                }

                // 2. Add a new named range
                int newIdx = names.Add("NewRange");
                Name newRange = names[newIdx];
                newRange.RefersTo = "=Sheet2!$A$1:$A$10";

                // 3. Delete an existing named range
                if (names["OldRange"] != null)
                {
                    names.Remove("OldRange");
                }

                // -----------------------------------------------------------------
                // Detect changes by comparing the current state with the original state
                // -----------------------------------------------------------------

                var logEntries = new List<string>();

                // Check for added or modified named ranges
                foreach (Name currentName in names)
                {
                    string nameText = currentName.Text;
                    string currentRef = currentName.RefersTo;

                    if (originalReferences.TryGetValue(nameText, out string oldRef))
                    {
                        // Modified?
                        if (!string.Equals(oldRef, currentRef, StringComparison.OrdinalIgnoreCase))
                        {
                            string entry = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - Modified: '{nameText}' " +
                                           $"Old Ref = {oldRef}, New Ref = {currentRef}";
                            logEntries.Add(entry);
                        }

                        // Remove processed entry to later identify deletions
                        originalReferences.Remove(nameText);
                    }
                    else
                    {
                        // Newly added range
                        string entry = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - Added: '{nameText}' " +
                                       $"Ref = {currentRef}";
                        logEntries.Add(entry);
                    }
                }

                // Remaining items are deleted named ranges
                foreach (var kvp in originalReferences)
                {
                    string entry = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - Deleted: '{kvp.Key}' " +
                                   $"Old Ref = {kvp.Value}";
                    logEntries.Add(entry);
                }

                // Append log entries to the log file
                File.AppendAllLines(logFilePath, logEntries);

                // Optionally, save the workbook if modifications need to be persisted
                // workbook.Save("output.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
