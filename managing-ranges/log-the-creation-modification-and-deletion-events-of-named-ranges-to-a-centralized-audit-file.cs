// Title: Audit Named Range Creation, Modification, and Deletion with Aspose.Cells for .NET
// Description: Demonstrates how to track every CREATE, MODIFY, and DELETE event of a named range in an Excel workbook using Aspose.Cells. The sample writes timestamped entries to a single audit file (NamedRangeAudit.log), shows the log after saving, and covers range definition, reference updates, and removal.
// Keywords: Aspose.Cells named range audit | C# log named range events | track named range changes .NET | centralized workbook audit file | timestamped Excel range logging
// Common Searches: how to log named range creation Aspose.Cells | track modifications to named ranges in .NET | audit deletion of Excel named ranges | centralized log for workbook changes C# | Aspose.Cells example for range audit
// Developer Intent: Record each creation, update, and removal of named ranges in an Excel workbook to a persistent log for compliance, debugging, or change‑tracking purposes.
// Use Cases: Create a named range (e.g., SalesData) and automatically write a CREATE entry with its RefersTo formula. | Change the RefersTo address of an existing named range and log a MODIFY entry that captures the new range. | Delete a named range from the workbook and generate a DELETE entry in the audit file.
// AI Prompts: Generate C# code that writes named‑range audit entries to a SQL Server table instead of a text file using Aspose.Cells. | Show how to add exponential‑backoff retry logic to the Log method when the audit file is locked. | Create a utility that scans multiple workbooks, aggregates all named‑range audit records, and exports a consolidated CSV report.

using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace NamedRangeAuditDemo
{
    // Demonstrates how to track every CREATE, MODIFY, and DELETE event of a named range in an Excel workbook using Aspose.Cells. The sample writes timestamped entries to a single audit file (NamedRangeAudit.log), shows the log after saving, and covers range definition, reference updates, and removal.
    class Program
    {
        // Path to the centralized audit file
        private const string AuditFilePath = "NamedRangeAudit.log";

        // Helper method to write audit entries with timestamp
        static void Log(string action, string name, string details = "")
        {
            try
            {
                string entry = $"{DateTime.UtcNow:O} | {action} | Name: {name}";
                if (!string.IsNullOrEmpty(details))
                    entry += $" | Details: {details}";
                entry += Environment.NewLine;

                File.AppendAllText(AuditFilePath, entry);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Logging failed: {ex.Message}");
            }
        }

        static void Main()
        {
            try
            {
                // Ensure the audit file starts fresh for this run
                if (File.Exists(AuditFilePath))
                    File.Delete(AuditFilePath);

                // ---------- Create a new workbook ----------
                Workbook workbook = new Workbook(); // create rule

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // ---------- Create a named range ----------
                // Define a range A1:B2 and assign a name
                AsposeRange range = sheet.Cells.CreateRange("A1", "B2");
                range.Name = "SalesData";

                // Add the name to the workbook's name collection
                int nameIndex = workbook.Worksheets.Names.Add(range.Name);
                Name namedRange = workbook.Worksheets.Names[nameIndex];
                // Set RefersTo using proper A1 notation
                namedRange.RefersTo = $"={sheet.Name}!$A$1:$B$2";
                // Log creation
                Log("CREATE", namedRange.Text, $"RefersTo = {namedRange.RefersTo}");

                // ---------- Modify the named range ----------
                // Change the reference to include an extra column (A1:C2)
                namedRange.RefersTo = $"={sheet.Name}!$A$1:$C$2";
                // Log modification
                Log("MODIFY", namedRange.Text, $"New RefersTo = {namedRange.RefersTo}");

                // ---------- Delete the named range ----------
                // Remove the name from the collection
                workbook.Worksheets.Names.Remove(namedRange.Text);
                // Log deletion
                Log("DELETE", namedRange.Text);

                // ---------- Save the workbook ----------
                string outputPath = "NamedRangeAuditDemo.xlsx";
                workbook.Save(outputPath); // save rule

                // Optional: display audit file content
                Console.WriteLine("Audit Log:");
                if (File.Exists(AuditFilePath))
                    Console.WriteLine(File.ReadAllText(AuditFilePath));
                else
                    Console.WriteLine("No audit log found.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
