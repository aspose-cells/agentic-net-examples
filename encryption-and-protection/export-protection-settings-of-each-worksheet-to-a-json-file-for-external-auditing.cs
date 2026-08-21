// Title: Export Excel Worksheet Protection Settings to JSON with Aspose.Cells for .NET (C#)
// Description: Creates a workbook, applies various protection options (including passwords and full protection), extracts each worksheet's protection flags and password hash, serializes the data into an indented JSON file (WorksheetProtectionAudit.json), and saves it for external auditing or compliance checks.
// Keywords: Aspose.Cells | C# | export worksheet protection | Excel protection JSON | worksheet password hash | serialize protection settings | audit Excel security | protect worksheet programmatically | compliance reporting Excel
// Common Searches: Aspose.Cells export worksheet protection to JSON | How to get Excel sheet protection settings with Aspose.Cells .NET | Retrieve password hash of protected worksheet using Aspose.Cells | Serialize Excel worksheet protection flags to JSON | Audit Excel sheet security with Aspose.Cells C#
// Developer Intent: Generate a JSON file that lists the protection configuration of every worksheet in a workbook for audit or compliance purposes.
// Use Cases: Produce compliance reports that show which sheets are password‑protected and which actions are allowed. | Compare protection configurations across multiple workbooks to enforce corporate security policies. | Log detailed protection data before sharing a workbook with partners, enabling later verification of sheet security.
// AI Prompts: Write C# code that reads WorksheetProtectionAudit.json and flags any sheet that allows row deletion without a password. | Show how to deserialize the exported JSON into objects and display a summary of each worksheet's protection options. | Explain how to extend the JSON export to include custom metadata such as the user who applied protection or the protection timestamp.

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Creates a workbook, applies various protection options (including passwords and full protection), extracts each worksheet's protection flags and password hash, serializes the data into an indented JSON file (WorksheetProtectionAudit.json), and saves it for external auditing or compliance checks.
    class ExportWorksheetProtection
    {
        static void Main()
        {
            try
            {
                // Create a workbook and add sample worksheets
                Workbook workbook = new Workbook();
                workbook.Worksheets.Add("Sheet2");
                workbook.Worksheets.Add("Sheet3");

                // Configure protection for each worksheet
                // Sheet1
                Protection p1 = workbook.Worksheets[0].Protection;
                p1.AllowEditingObject = false;
                p1.AllowEditingContent = true;
                p1.Password = "pwd1";

                // Sheet2
                Protection p2 = workbook.Worksheets[1].Protection;
                p2.AllowDeletingRow = true;
                p2.AllowInsertingColumn = true;
                p2.Password = "pwd2";

                // Sheet3 - protect all without password
                workbook.Worksheets[2].Protect(ProtectionType.All);

                // Collect protection settings from each worksheet
                var sheetInfos = new List<object>();
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    Protection p = sheet.Protection;
                    sheetInfos.Add(new
                    {
                        SheetName = sheet.Name,
                        IsProtectedWithPassword = p.IsProtectedWithPassword,
                        // GetPasswordHash returns int; use nullable int for JSON serialization
                        PasswordHash = p.IsProtectedWithPassword ? (int?)p.GetPasswordHash() : null,
                        p.AllowDeletingColumn,
                        p.AllowDeletingRow,
                        p.AllowEditingContent,
                        p.AllowEditingObject,
                        p.AllowEditingScenario,
                        p.AllowFiltering,
                        p.AllowFormattingCell,
                        p.AllowFormattingColumn,
                        p.AllowFormattingRow,
                        p.AllowInsertingColumn,
                        p.AllowInsertingHyperlink,
                        p.AllowInsertingRow,
                        p.AllowSelectingLockedCell,
                        p.AllowSelectingUnlockedCell,
                        p.AllowSorting,
                        p.AllowUsingPivotTable
                    });
                }

                // Serialize the collected data to JSON
                var jsonOptions = new JsonSerializerOptions { WriteIndented = true };
                string json = JsonSerializer.Serialize(sheetInfos, jsonOptions);

                // Write JSON to a file for auditing
                string outputPath = "WorksheetProtectionAudit.json";
                File.WriteAllText(outputPath, json);

                Console.WriteLine($"Protection settings exported to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
