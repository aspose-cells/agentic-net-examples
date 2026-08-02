// Title: Batch Update Worksheet TabId and Generate CSV Summary with Aspose.Cells for .NET
// Description: Scans a folder for Excel workbooks, loads each file with Aspose.Cells, increments every worksheet's TabId, saves the changes, and writes a CSV log (file, sheet, old TabId, new TabId) to the same directory.
// Keywords: Aspose.Cells TabId update | C# batch Excel processing | increment worksheet TabId | generate CSV change log | process multiple workbooks | bulk worksheet property edit
// Common Searches: update TabId for all worksheets in a folder using Aspose.Cells | C# create CSV report of worksheet TabId changes | batch modify Excel TabId property .NET | automate TabId renumbering across many workbooks | Aspose.Cells example for bulk worksheet updates
// Developer Intent: Iterate through every Excel file in a directory, change each worksheet's TabId, persist the workbook, and record old and new TabId values in a summary CSV.
// Use Cases: Standardize tab ordering in a collection of financial models before distribution. | Maintain an audit trail of TabId adjustments when migrating legacy spreadsheets. | Quickly verify bulk TabId changes by reviewing a generated CSV report.
// AI Prompts: Generate C# code with Aspose.Cells that increments the TabId of each worksheet in all Excel files of a given folder and outputs a CSV log. | Explain how to handle file‑access errors and avoid workbook locks during bulk TabId updates with Aspose.Cells. | Suggest ways to customize the TabId logic (e.g., set based on sheet index or naming pattern) while still producing a change summary CSV.

using System;
using System.IO;
using System.Text;
using System.Linq;
using Aspose.Cells;

namespace ExcelTabIdUpdater
{
    // Scans a folder for Excel workbooks, loads each file with Aspose.Cells, increments every worksheet's TabId, saves the changes, and writes a CSV log (file, sheet, old TabId, new TabId) to the same directory.
    class UpdateTabIdsAndSummarize
    {
        static void Main()
        {
            // Directory containing the Excel files
            string sourceDirectory = @"C:\ExcelFiles";

            // Verify that the source directory exists
            if (!Directory.Exists(sourceDirectory))
            {
                Console.WriteLine($"Source directory not found: {sourceDirectory}");
                return;
            }

            // Path for the summary CSV file
            string summaryCsvPath = Path.Combine(sourceDirectory, "TabIdSummary.csv");

            // Prepare CSV header
            StringBuilder csvBuilder = new StringBuilder();
            csvBuilder.AppendLine("FileName,WorksheetName,OldTabId,NewTabId");

            try
            {
                // Get all Excel files in the directory
                var excelFiles = Directory.GetFiles(sourceDirectory, "*.*", SearchOption.TopDirectoryOnly)
                    .Where(f => f.EndsWith(".xls", StringComparison.OrdinalIgnoreCase) ||
                                f.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) ||
                                f.EndsWith(".xlsm", StringComparison.OrdinalIgnoreCase) ||
                                f.EndsWith(".xlsb", StringComparison.OrdinalIgnoreCase));

                foreach (string filePath in excelFiles)
                {
                    // Ensure the file exists before attempting to load
                    if (!File.Exists(filePath))
                    {
                        Console.WriteLine($"File not found, skipping: {filePath}");
                        continue;
                    }

                    // Load the workbook
                    Workbook workbook = new Workbook(filePath);

                    // Iterate through all worksheets and update TabId
                    foreach (Worksheet sheet in workbook.Worksheets)
                    {
                        int oldTabId = sheet.TabId;
                        int newTabId = oldTabId + 1; // Example update: increment by 1
                        sheet.TabId = newTabId;

                        // Record the change in the CSV summary
                        csvBuilder.AppendLine($"{Path.GetFileName(filePath)},{sheet.Name},{oldTabId},{newTabId}");
                    }

                    // Save the modified workbook back to the same file
                    workbook.Save(filePath);
                }

                // Write the summary CSV to disk
                File.WriteAllText(summaryCsvPath, csvBuilder.ToString());

                Console.WriteLine("TabId update completed. Summary saved to: " + summaryCsvPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred during processing:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
