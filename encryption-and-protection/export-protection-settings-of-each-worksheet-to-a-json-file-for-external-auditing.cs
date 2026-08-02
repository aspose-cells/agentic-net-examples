// Title: Export Excel Worksheet Protection Settings to JSON with Aspose.Cells for .NET
// Description: Load an Excel workbook using Aspose.Cells, read each worksheet's protection flags (including password presence and allowed actions), serialize the data to a formatted JSON file, and save it for compliance auditing or security reviews.
// Keywords: Aspose.Cells worksheet protection export | C# Excel protection JSON | audit Excel sheet security | retrieve worksheet protection flags .NET | serialize Excel protection settings | Excel password audit Aspose | Excel sheet protection report | Aspose.Cells security compliance | export workbook protection to JSON | list worksheet protection properties
// Common Searches: Aspose.Cells export worksheet protection to JSON | C# get Excel sheet protection details | How to audit Excel worksheet passwords with Aspose | Serialize worksheet protection flags in .NET | Create JSON report of Excel sheet security settings
// Developer Intent: Extract protection configuration of every worksheet and write it to a JSON file for external auditing.
// Use Cases: Generate compliance reports that list protection flags and password status for each sheet. | Detect security changes by comparing protection JSON from two workbook versions. | Log worksheet protection settings in CI/CD pipelines for automatically generated reports.
// AI Prompts: Write C# code using Aspose.Cells to read all worksheet protection properties and output them as an indented JSON file. | Modify the example so the password field is excluded from the JSON output for security compliance. | Create a reusable method that returns a dictionary mapping worksheet names to their protection status and allowed actions.

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

// Load an Excel workbook using Aspose.Cells, read each worksheet's protection flags (including password presence and allowed actions), serialize the data to a formatted JSON file, and save it for compliance auditing or security reviews.
class ExportWorksheetProtection
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // List to hold protection info for each worksheet
        var protectionInfoList = new List<object>();

        // Iterate through all worksheets in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            Protection protection = sheet.Protection;

            // Capture relevant protection settings
            var sheetInfo = new
            {
                SheetName = sheet.Name,
                IsProtectedWithPassword = protection.IsProtectedWithPassword,
                Password = protection.Password, // empty if not set
                AllowDeletingColumn = protection.AllowDeletingColumn,
                AllowDeletingRow = protection.AllowDeletingRow,
                AllowEditingContent = protection.AllowEditingContent,
                AllowEditingObject = protection.AllowEditingObject,
                AllowEditingScenario = protection.AllowEditingScenario,
                AllowFiltering = protection.AllowFiltering,
                AllowFormattingCell = protection.AllowFormattingCell,
                AllowFormattingColumn = protection.AllowFormattingColumn,
                AllowFormattingRow = protection.AllowFormattingRow,
                AllowInsertingColumn = protection.AllowInsertingColumn,
                AllowInsertingHyperlink = protection.AllowInsertingHyperlink,
                AllowInsertingRow = protection.AllowInsertingRow,
                AllowSelectingLockedCell = protection.AllowSelectingLockedCell,
                AllowSelectingUnlockedCell = protection.AllowSelectingUnlockedCell,
                AllowSorting = protection.AllowSorting,
                AllowUsingPivotTable = protection.AllowUsingPivotTable
            };

            protectionInfoList.Add(sheetInfo);
        }

        // Serialize the list to a formatted JSON string
        string jsonOutput = JsonSerializer.Serialize(protectionInfoList, new JsonSerializerOptions { WriteIndented = true });

        // Write the JSON to a file for external auditing
        File.WriteAllText("WorksheetProtectionAudit.json", jsonOutput);

        Console.WriteLine("Worksheet protection settings have been exported to WorksheetProtectionAudit.json");
    }
}
