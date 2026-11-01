// Title: C# code to log added, modified, and deleted named ranges with timestamps using Aspose.Cells
// Common Searches: how to audit named range changes in an Excel file with Aspose.Cells C# | C# write timestamped log of added and removed named ranges | track modifications to named ranges and export to text file using Aspose.Cells | compare original and current named ranges in .NET and generate change report | Aspose.Cells example for logging named range updates with old and new references
// Tags: named range change tracking Aspose.Cells | timestamped named range log C# | detect added deleted named ranges .NET | export named range modifications to text file | original vs current named range comparison Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// This C# example loads an Excel workbook with Aspose.Cells, captures the initial set of named ranges, then after any additions, updates, or deletions it compares the current named ranges to the original collection and writes a timestamped log file that records added, modified (showing old and new references) and deleted ranges, before saving the workbook.
class NamedRangeLogger
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        string workbookPath = "InputWorkbook.xlsx";
        Workbook workbook = new Workbook(workbookPath);

        // Capture the initial state of all named ranges
        Dictionary<string, string> originalRanges = new Dictionary<string, string>();
        foreach (Name name in workbook.Worksheets.Names)
        {
            // Store the name and its reference (e.g., Sheet1!$A$1:$B$2)
            originalRanges[name.Text] = name.RefersTo;
        }

        // -----------------------------------------------------------------
        // Perform modifications to named ranges here.
        // Example modifications (replace with actual logic):
        //   - Add a new named range
        //   - Change the reference of an existing named range
        //   - Delete a named range
        // -----------------------------------------------------------------
        // Example: Add a new named range
        // workbook.Worksheets.Names.Add("NewRange", "Sheet1!$C$1:$D$5");

        // Example: Change an existing named range
        // Name existing = workbook.Worksheets.Names["ExistingRange"];
        // existing.RefersTo = "Sheet2!$A$1:$A$10";

        // Example: Delete a named range
        // workbook.Worksheets.Names.Remove("ObsoleteRange");

        // -----------------------------------------------------------------
        // End of modifications
        // -----------------------------------------------------------------

        // Prepare the log file
        string logFilePath = "NamedRangeChanges.log";
        using (StreamWriter logWriter = new StreamWriter(logFilePath, true))
        {
            // Timestamp for this logging session
            string timeStamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            logWriter.WriteLine($"--- Modification Log: {timeStamp} ---");

            // Capture the current state of named ranges
            Dictionary<string, string> currentRanges = new Dictionary<string, string>();
            foreach (Name name in workbook.Worksheets.Names)
            {
                currentRanges[name.Text] = name.RefersTo;
            }

            // Detect added or modified named ranges
            foreach (var kvp in currentRanges)
            {
                string name = kvp.Key;
                string newRef = kvp.Value;

                if (originalRanges.TryGetValue(name, out string oldRef))
                {
                    if (!string.Equals(oldRef, newRef, StringComparison.OrdinalIgnoreCase))
                    {
                        // Modified range
                        logWriter.WriteLine($"[Modified] Name: {name}");
                        logWriter.WriteLine($"    Old Reference: {oldRef}");
                        logWriter.WriteLine($"    New Reference: {newRef}");
                    }
                    // Remove from original dictionary to simplify later detection of deletions
                    originalRanges.Remove(name);
                }
                else
                {
                    // Added range
                    logWriter.WriteLine($"[Added] Name: {name}");
                    logWriter.WriteLine($"    Reference: {newRef}");
                }
            }

            // Remaining entries in originalRanges are deleted ranges
            foreach (var kvp in originalRanges)
            {
                logWriter.WriteLine($"[Deleted] Name: {kvp.Key}");
                logWriter.WriteLine($"    Last Reference: {kvp.Value}");
            }

            logWriter.WriteLine(); // Blank line for readability
        }

        // Save the workbook if needed (replace with your desired output path)
        string outputPath = "OutputWorkbook.xlsx";
        workbook.Save(outputPath);
    }
}
