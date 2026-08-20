// Title: Track Named Range Modifications with Timestamps via Aspose.Cells Revision Logs in C#
// Description: The sample creates a shared workbook, defines a workbook‑level named range, updates its reference several times while saving after each change, then reads the workbook’s RevisionLogs to extract DefinedName revisions. Each revision is written to a plain‑text file with the current timestamp, the name, the old formula and the new formula.
// Keywords: Aspose.Cells | C# revision log | named range audit | track defined name changes | shared workbook | Excel revision tracking | log named range modifications | timestamped change history
// Common Searches: Aspose.Cells log named range changes | revision logs defined name C# | write named range history to text file | track named range revisions in .NET | audit Excel named ranges with Aspose
// Developer Intent: Automatically capture every change to workbook named ranges, recording the time, previous reference, and new reference in a readable log file.
// Use Cases: Compliance reporting for financial models that rely on dynamic named ranges | Debugging automated spreadsheet updates by reviewing a chronological change log | Generating audit trails for shared Excel workbooks in collaborative environments
// AI Prompts: Write C# code that reads Aspose.Cells RevisionLogs, filters for RevisionDefinedName entries, and exports the data to a CSV with columns: Timestamp, Name, OldFormula, NewFormula. | Explain how to enable revision tracking for a workbook, modify a named range, and retrieve its revision history using Aspose.Cells. | Provide a concise guide to log named range changes without overwriting existing log entries, ensuring each entry includes a timestamp.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Revisions;

// The sample creates a shared workbook, defines a workbook‑level named range, updates its reference several times while saving after each change, then reads the workbook’s RevisionLogs to extract DefinedName revisions. Each revision is written to a plain‑text file with the current timestamp, the name, the old formula and the new formula.
class NamedRangeLogger
{
    static void Main()
    {
        // Paths for the workbook and the log file
        string workbookPath = "NamedRangeLogDemo.xlsx";
        string logPath = "NamedRangeChanges.txt";

        try
        {
            // -------------------------------------------------
            // Create a workbook and enable shared mode to track revisions
            // -------------------------------------------------
            Workbook wb = new Workbook();
            wb.Settings.Shared = true; // Enable shared workbook (required for revision tracking)

            Worksheet ws = wb.Worksheets[0];
            ws.Name = "Sheet1";

            // -------------------------------------------------
            // Create an initial named range (workbook‑level name)
            // -------------------------------------------------
            int nameIdx = wb.Worksheets.Names.Add("MyRange");
            Name namedRange = wb.Worksheets.Names[nameIdx];
            namedRange.RefersTo = "='Sheet1'!$A$1:$A$3";

            // Save the first version (creates the initial revision entry)
            wb.Save(workbookPath);

            // -------------------------------------------------
            // Modify the named range multiple times, saving after each change
            // -------------------------------------------------
            namedRange.RefersTo = "='Sheet1'!$A$1:$A$4";
            wb.Save(workbookPath);

            namedRange.RefersTo = "='Sheet1'!$B$1:$B$4";
            wb.Save(workbookPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error during workbook creation/modification: {ex.Message}");
            return;
        }

        // -------------------------------------------------
        // Reopen the workbook to read revision logs
        // -------------------------------------------------
        if (!File.Exists(workbookPath))
        {
            Console.WriteLine($"Workbook file not found: {workbookPath}");
            return;
        }

        try
        {
            Workbook revWb = new Workbook(workbookPath);

            // Open a StreamWriter to write the log entries
            using (StreamWriter writer = new StreamWriter(logPath, false))
            {
                // Iterate through all revision logs in the workbook
                foreach (RevisionLog log in revWb.Worksheets.RevisionLogs)
                {
                    // Iterate through each revision in the log
                    foreach (Revision rev in log.Revisions)
                    {
                        // We're interested only in defined name revisions
                        if (rev.Type == RevisionType.DefinedName)
                        {
                            RevisionDefinedName nameRev = (RevisionDefinedName)rev;

                            // Timestamp for when the log entry is written
                            string timeStamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

                            // Write details to the log file
                            writer.WriteLine($"{timeStamp} - Defined Name: {nameRev.Text}");
                            writer.WriteLine($"    Old Formula: {nameRev.OldFormula}");
                            writer.WriteLine($"    New Formula: {nameRev.NewFormula}");
                        }
                    }
                }
            }

            // Optional: display the generated log on console
            if (File.Exists(logPath))
            {
                Console.WriteLine(File.ReadAllText(logPath));
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error during revision reading/logging: {ex.Message}");
        }
    }
}
