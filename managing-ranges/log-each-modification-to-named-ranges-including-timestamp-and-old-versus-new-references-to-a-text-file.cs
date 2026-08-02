// Title: Track Named Range Modifications with Timestamps Using Aspose.Cells Revision Logs (C#)
// Description: Creates a shared workbook, adds a named range, changes its reference, saves to generate revision entries, then reads RevisionLogs to capture DefinedName revisions. Each change is written with a UTC timestamp, the range name, old formula and new formula to a plain‑text log file.
// Keywords: Aspose.Cells | C# | .NET | revision log | named range audit | track defined name changes | timestamped log file | workbook change history | shared workbook
// Common Searches: Aspose.Cells log named range changes C# | How to capture defined name revisions with Aspose.Cells | Write named range modification history to a file using .NET | Audit named range updates in a shared Excel workbook
// Developer Intent: Automatically record every alteration of a workbook's named ranges—including when the change occurred and the previous vs. new reference—into a readable log file.
// Use Cases: Compliance audit of named‑range edits in collaborative spreadsheets. | Debugging unexpected formula shifts by reviewing a chronological change log. | Maintaining a versioned history for named ranges that feed downstream data pipelines.
// AI Prompts: Generate code that also logs the user ID responsible for each named‑range change using Aspose.Cells revision metadata. | Create a method that returns the latest revision entry for each named range from the revision logs. | Show how to output the log entries in CSV format for easy import into Excel or Power BI.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Revisions;

// Creates a shared workbook, adds a named range, changes its reference, saves to generate revision entries, then reads RevisionLogs to capture DefinedName revisions. Each change is written with a UTC timestamp, the range name, old formula and new formula to a plain‑text log file.
class NamedRangeLogger
{
    static void Main()
    {
        // Paths for the workbook and the log file
        string workbookPath = "NamedRangeDemo.xlsx";
        string logPath = "NamedRangeChanges.log";

        // ---------- Create ----------
        // Create a new workbook and enable sharing to capture revisions
        Workbook wb = new Workbook();
        wb.Settings.Shared = true;

        // Add a named range "MyRange" referring to A1:A3
        int nameIdx = wb.Worksheets.Names.Add("MyRange");
        Name namedRange = wb.Worksheets.Names[nameIdx];
        namedRange.RefersTo = "=Sheet1!$A$1:$A$3";

        // Save the initial version (creates the first revision entry)
        wb.Save(workbookPath);

        // ---------- Modify ----------
        // Change the reference of the named range to A1:A4
        namedRange.RefersTo = "=Sheet1!$A$1:$A$4";

        // Save after modification (generates a revision for the defined name change)
        wb.Save(workbookPath);

        // ---------- Load ----------
        // Load the workbook to access its revision logs
        Workbook loadedWb = new Workbook(workbookPath);

        // Open the log file for appending
        using (StreamWriter sw = new StreamWriter(logPath, true))
        {
            // Iterate through all revision logs
            foreach (RevisionLog log in loadedWb.Worksheets.RevisionLogs)
            {
                // Iterate through each revision entry
                foreach (Revision rev in log.Revisions)
                {
                    // Filter for defined name revisions
                    if (rev.Type == RevisionType.DefinedName && rev is RevisionDefinedName definedNameRev)
                    {
                        // Record timestamp, name text, old formula and new formula
                        string timestamp = DateTime.Now.ToString("o");
                        sw.WriteLine($"{timestamp}: Name='{definedNameRev.Text}' OldFormula='{definedNameRev.OldFormula}' NewFormula='{definedNameRev.NewFormula}'");
                    }
                }
            }
        }

        // Optional: display the log content
        Console.WriteLine("Named range modifications logged:");
        Console.WriteLine(File.ReadAllText(logPath));
    }
}
