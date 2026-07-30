// Title: Export Excel worksheet protection settings to JSON using Aspose.Cells for .NET
// Description: A C# example that loads an Excel workbook with Aspose.Cells, reads each worksheet's Protection object (password status, hash, and all Allow* flags), serializes the data into a formatted JSON file, and saves it for external audit or compliance reporting.
// Keywords: Aspose.Cells | C# | export worksheet protection | Excel protection JSON | password hash extraction | audit Excel security | .NET workbook protection | serialize protection settings
// Common Searches: How to export Excel sheet protection to JSON with Aspose.Cells | Retrieve worksheet password hash in C# | Aspose.Cells export protection flags for audit | Create JSON report of Excel worksheet security | C# code to list Allow* protection options per sheet
// Developer Intent: Generate a JSON audit file that captures the protection configuration of every worksheet in an Excel workbook.
// Use Cases: Produce compliance reports showing which sheets are password‑protected and which editing actions are allowed. | Compare protection policies across multiple workbooks to detect security gaps. | Feed worksheet protection data into governance tools or SIEM systems for continuous monitoring.
// AI Prompts: Write a C# routine that reads the generated WorksheetProtectionAudit.json and flags any sheet that permits row deletion without a password. | Create a PowerShell script to upload the JSON audit file to SharePoint and record the upload timestamp. | Suggest a method to mask the PasswordHash field in the JSON output while keeping other protection flags readable.

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

namespace AsposeCellsProtectionExport
{
    // A C# example that loads an Excel workbook with Aspose.Cells, reads each worksheet's Protection object (password status, hash, and all Allow* flags), serializes the data into a formatted JSON file, and saves it for external audit or compliance reporting.
    class Program
    {
        static void Main()
        {
            // Load the workbook that contains protected worksheets
            string workbookPath = "protected.xlsx";
            Workbook workbook = new Workbook(workbookPath);

            // Prepare a list to hold protection information for each worksheet
            var worksheetsProtectionInfo = new List<Dictionary<string, object>>();

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Access the Protection object of the current worksheet
                Protection protection = sheet.Protection;

                // Collect relevant protection settings into a dictionary
                var sheetInfo = new Dictionary<string, object>
                {
                    { "WorksheetName", sheet.Name },
                    { "IsProtectedWithPassword", protection.IsProtectedWithPassword },
                    { "PasswordHash", protection.GetPasswordHash() },
                    { "AllowDeletingColumn", protection.AllowDeletingColumn },
                    { "AllowDeletingRow", protection.AllowDeletingRow },
                    { "AllowEditingContent", protection.AllowEditingContent },
                    { "AllowEditingObject", protection.AllowEditingObject },
                    { "AllowEditingScenario", protection.AllowEditingScenario },
                    { "AllowFiltering", protection.AllowFiltering },
                    { "AllowFormattingCell", protection.AllowFormattingCell },
                    { "AllowFormattingColumn", protection.AllowFormattingColumn },
                    { "AllowFormattingRow", protection.AllowFormattingRow },
                    { "AllowInsertingColumn", protection.AllowInsertingColumn },
                    { "AllowInsertingHyperlink", protection.AllowInsertingHyperlink },
                    { "AllowInsertingRow", protection.AllowInsertingRow },
                    { "AllowSelectingLockedCell", protection.AllowSelectingLockedCell },
                    { "AllowSelectingUnlockedCell", protection.AllowSelectingUnlockedCell },
                    { "AllowSorting", protection.AllowSorting },
                    { "AllowUsingPivotTable", protection.AllowUsingPivotTable }
                };

                worksheetsProtectionInfo.Add(sheetInfo);
            }

            // Serialize the collected information to a formatted JSON string
            var jsonOptions = new JsonSerializerOptions { WriteIndented = true };
            string jsonOutput = JsonSerializer.Serialize(worksheetsProtectionInfo, jsonOptions);

            // Save the JSON to a file for external auditing
            string outputPath = "WorksheetProtectionAudit.json";
            File.WriteAllText(outputPath, jsonOutput);

            Console.WriteLine($"Protection settings exported to: {outputPath}");
        }
    }
}
