// Title: Update the 'TabId' custom document property in all .xlsx workbooks in a folder and generate a CSV log of changes using Aspose.Cells for .NET
// AI Prompts: Write a C# console program that enumerates every .xlsx file in a given directory, reads the existing 'TabId' custom document property with Aspose.Cells, replaces it with a newly generated GUID (or adds it if missing), saves each workbook, and records the file name, previous TabId, and new GUID to a UTF‑8 CSV file. | Generate C# code that uses Aspose.Cells to batch‑process Excel files, update or insert a custom property named 'TabId' with a unique identifier, persist each workbook, and output a summary CSV containing columns for FileName, OldTabId, and NewTabId.
// Common Searches: how to batch replace a custom document property in multiple Excel files using Aspose.Cells C# | c# program to add GUID to TabId property of all .xlsx workbooks in a folder | generate a CSV report of old and new custom property values after updating Excel files with Aspose.Cells | Aspose.Cells iterate through directory and modify custom properties in .xlsx files
// Tags: Aspose.Cells batch update custom document property | C# modify TabId property in .xlsx workbooks | generate CSV change log for Excel custom properties | GUID assignment to Excel custom property using Aspose.Cells | process multiple Excel files programmatically .NET

using System;
using System.IO;
using System.Text;
using Aspose.Cells;
using System.Collections.Generic;

namespace ExcelTabIdUpdater
{
    // The program scans a specified folder for .xlsx files, loads each workbook with Aspose.Cells, reads the existing 'TabId' custom document property, replaces or adds it with a new GUID, saves the workbook, and writes a UTF‑8 CSV summarizing FileName, OldTabId, and NewTabId.
    class Program
    {
        static void Main(string[] args)
        {
            // Directory containing the Excel files
            string sourceDirectory = @"C:\ExcelFiles";
            // Path for the summary CSV file
            string summaryCsvPath = Path.Combine(sourceDirectory, "TabIdChangesSummary.csv");

            // Prepare a StringBuilder for CSV content
            StringBuilder csvBuilder = new StringBuilder();
            csvBuilder.AppendLine("FileName,OldTabId,NewTabId");

            // Get all .xlsx files in the directory (including subfolders if needed)
            string[] excelFiles = Directory.GetFiles(sourceDirectory, "*.xlsx", SearchOption.TopDirectoryOnly);

            foreach (string filePath in excelFiles)
            {
                // Load the workbook using Aspose.Cells
                Workbook workbook = new Workbook(filePath);

                // Access the custom document properties collection
                var customProps = workbook.CustomDocumentProperties;

                // Retrieve the existing TabId (if any)
                string oldTabId = string.Empty;
                if (customProps.Contains("TabId"))
                {
                    oldTabId = customProps["TabId"].Value?.ToString() ?? string.Empty;
                }

                // Generate a new TabId (using a GUID for uniqueness)
                string newTabId = Guid.NewGuid().ToString();

                // Update or add the TabId property
                if (customProps.Contains("TabId"))
                {
                    customProps["TabId"].Value = newTabId;
                }
                else
                {
                    customProps.Add("TabId", newTabId);
                }

                // Save the workbook back to the same file
                workbook.Save(filePath);

                // Append change information to the CSV
                string fileName = Path.GetFileName(filePath);
                csvBuilder.AppendLine($"{fileName},{oldTabId},{newTabId}");
            }

            // Write the summary CSV to disk
            File.WriteAllText(summaryCsvPath, csvBuilder.ToString(), Encoding.UTF8);
        }
    }
}
