// Title: Batch Update Worksheet TabIds and Export Change Log to CSV with Aspose.Cells for .NET
// Description: Scans a folder for .xlsx files, loads each workbook with Aspose.Cells, iterates all worksheets, records the original TabId, assigns a new TabId based on the worksheet index, saves the workbook, and writes a CSV summary (FileName, WorksheetName, OldTabId, NewTabId).
// Keywords: Aspose.Cells TabId update | C# batch Excel processing | modify worksheet tab order programmatically | generate CSV change log Aspose | process multiple workbooks .NET | Excel TabId automation | bulk worksheet property update
// Common Searches: How to change TabId for all worksheets in a folder using Aspose.Cells | C# code to batch update Excel worksheet TabIds and create a CSV report | Aspose.Cells example for enumerating worksheets and saving changes | Generate a change log of worksheet TabId modifications in .NET
// Developer Intent: Iterate through every .xlsx file in a directory, set each worksheet’s TabId to its index + 1, and log old and new values to a CSV file.
// Use Cases: Standardize tab order across a collection of reports before distribution | Create an audit trail of worksheet identifier changes for compliance | Integrate automatic TabId normalization into a CI/CD pipeline for generated workbooks
// AI Prompts: Write C# code with Aspose.Cells that updates each worksheet TabId to its position and outputs a CSV log of the changes. | Provide a version of the batch TabId updater that includes robust error handling for locked or corrupted files. | Show how to replace the index‑based TabId assignment with a custom mapping supplied from a JSON configuration.

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

// Scans a folder for .xlsx files, loads each workbook with Aspose.Cells, iterates all worksheets, records the original TabId, assigns a new TabId based on the worksheet index, saves the workbook, and writes a CSV summary (FileName, WorksheetName, OldTabId, NewTabId).
class Program
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

        // StringBuilder to accumulate CSV rows
        var csvBuilder = new StringBuilder();
        csvBuilder.AppendLine("FileName,WorksheetName,OldTabId,NewTabId");

        // Process each .xlsx file in the directory
        foreach (string excelFilePath in Directory.GetFiles(sourceDirectory, "*.xlsx"))
        {
            // Ensure the file exists before attempting to load
            if (!File.Exists(excelFilePath))
            {
                Console.WriteLine($"File not found (skipped): {excelFilePath}");
                continue;
            }

            try
            {
                // Load the workbook
                var workbook = new Workbook(excelFilePath);

                // Iterate through all worksheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Capture the original TabId
                    int oldTabId = sheet.TabId;

                    // Example update: set TabId to (worksheet index + 1)
                    int newTabId = sheet.Index + 1;
                    sheet.TabId = newTabId;

                    // Record the change in the CSV
                    csvBuilder.AppendLine($"{Path.GetFileName(excelFilePath)},{sheet.Name},{oldTabId},{newTabId}");
                }

                // Save the modified workbook
                workbook.Save(excelFilePath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing file '{excelFilePath}': {ex.Message}");
            }
        }

        try
        {
            // Write the summary CSV to disk
            File.WriteAllText(summaryCsvPath, csvBuilder.ToString());
            Console.WriteLine($"Summary CSV written to: {summaryCsvPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to write summary CSV: {ex.Message}");
        }
    }
}
