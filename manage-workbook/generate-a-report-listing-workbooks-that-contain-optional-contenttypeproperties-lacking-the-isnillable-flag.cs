// Title: C# – Generate CSV Report of Excel Workbooks with Non‑Nillable ContentTypeProperties using Aspose.Cells
// Description: A console app that scans a folder for .xls, .xlsx, .xlsm and .xlsb files, loads each workbook with Aspose.Cells, examines its ContentTypePropertyCollection, and records every property where the IsNillable flag is false or undefined. Results are saved to a CSV file with workbook name, property name and IsNillable value, and loading errors are logged.
// Keywords: Aspose.Cells ContentTypeProperty IsNillable | C# generate CSV report Excel workbooks | list non‑nillable content type properties | Aspose.Cells iterate ContentTypeProperties | detect optional ContentTypeProperty flag | Excel template compliance check | Aspose.Cells .NET CSV export
// Common Searches: Aspose.Cells code to list ContentTypeProperties without IsNillable | C# generate CSV of Excel files missing nillable flag | how to audit workbooks for non‑nillable content type properties | scan folder for Excel files and report IsNillable false | Aspose.Cells ContentTypePropertyCollection example
// Developer Intent: Create a CSV file that lists every workbook containing ContentTypeProperties where IsNillable is false or not set, including error handling for unreadable files.
// Use Cases: Validate that Excel templates mark optional content type properties as nillable before publishing. | Audit a repository of workbooks to prevent data loss during XML serialization. | Provide administrators with a quick‑look report to update non‑nillable properties.
// AI Prompts: Write C# code with Aspose.Cells to scan a directory of Excel files and output a CSV of ContentTypeProperties where IsNillable is false. | Improve the program to log errors to a separate file and enable recursive folder search. | Extend the report to include each property's DataType, default value, and whether it is required.

using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Properties;

namespace ContentTypePropertyReport
{
    // A console app that scans a folder for .xls, .xlsx, .xlsm and .xlsb files, loads each workbook with Aspose.Cells, examines its ContentTypePropertyCollection, and records every property where the IsNillable flag is false or undefined. Results are saved to a CSV file with workbook name, property name and IsNillable value, and loading errors are logged.
    class Program
    {
        // Entry point
        static void Main(string[] args)
        {
            // Folder containing the workbooks to analyze
            string folderPath = @"C:\Workbooks";

            // Output report file (CSV format)
            string reportPath = Path.Combine(folderPath, "ContentTypePropertiesReport.csv");

            // Prepare a list to hold report lines
            List<string> reportLines = new List<string>();
            // Header line
            reportLines.Add("WorkbookFile,PropertyName,IsNillable");

            // Get all Excel files in the folder (including subfolders if needed)
            string[] workbookFiles = Directory.GetFiles(folderPath, "*.*", SearchOption.TopDirectoryOnly);
            foreach (string file in workbookFiles)
            {
                // Filter supported Excel extensions
                string ext = Path.GetExtension(file).ToLowerInvariant();
                if (ext != ".xls" && ext != ".xlsx" && ext != ".xlsm" && ext != ".xlsb")
                    continue;

                try
                {
                    // Load the workbook
                    Workbook workbook = new Workbook(file);

                    // Iterate through all ContentTypeProperties
                    ContentTypePropertyCollection props = workbook.ContentTypeProperties;
                    for (int i = 0; i < props.Count; i++)
                    {
                        ContentTypeProperty prop = props[i];
                        // If IsNillable is false (or not set), record it
                        if (!prop.IsNillable)
                        {
                            string line = $"{Path.GetFileName(file)},{prop.Name},{prop.IsNillable}";
                            reportLines.Add(line);
                        }
                    }

                    // Dispose workbook (optional, as it implements IDisposable)
                    workbook.Dispose();
                }
                catch (Exception ex)
                {
                    // In case a file cannot be processed, write an error line
                    string errorLine = $"{Path.GetFileName(file)},Error loading workbook,{ex.Message}";
                    reportLines.Add(errorLine);
                }
            }

            // Write the report to the CSV file
            File.WriteAllLines(reportPath, reportLines);

            // Also output a summary to the console
            Console.WriteLine($"Report generated: {reportPath}");
            Console.WriteLine($"Total entries: {reportLines.Count - 1}");
        }
    }
}
