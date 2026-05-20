using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Properties;

class ContentTypePropertiesReport
{
    static void Main()
    {
        try
        {
            // Folder containing the workbooks to be inspected
            string folderPath = @"C:\Workbooks";

            // Output CSV file for the report
            string reportPath = @"C:\Report\ContentTypePropertiesReport.csv";

            // Prepare a list to hold report lines (CSV header + data)
            List<string> reportLines = new List<string>
            {
                "Workbook,PropertyName,PropertyType,PropertyValue,IsNillable"
            };

            // Get all Excel files in the folder
            string[] workbookFiles = Directory.GetFiles(folderPath, "*.xlsx", SearchOption.TopDirectoryOnly);

            foreach (string filePath in workbookFiles)
            {
                // Ensure the file still exists before attempting to load
                if (!File.Exists(filePath))
                    continue;

                try
                {
                    // Load the workbook (no password; if required, the exception will be caught)
                    using (Workbook workbook = new Workbook(filePath))
                    {
                        // Access the collection of ContentTypeProperty objects
                        ContentTypePropertyCollection ctProps = workbook.ContentTypeProperties;

                        // Iterate through each property
                        for (int i = 0; i < ctProps.Count; i++)
                        {
                            ContentTypeProperty prop = ctProps[i];

                            // If IsNillable is false, consider it "lacking the IsNillable flag"
                            if (!prop.IsNillable)
                            {
                                string line = string.Format(
                                    "\"{0}\",\"{1}\",\"{2}\",\"{3}\",{4}",
                                    Path.GetFileName(filePath),
                                    prop.Name,
                                    prop.Type,
                                    prop.Value?.Replace("\"", "\"\""),
                                    prop.IsNillable);

                                reportLines.Add(line);
                            }
                        }
                    }
                }
                catch (CellsException ex)
                {
                    // Detect password‑protected files via message content
                    if (ex.Message != null && ex.Message.IndexOf("password", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        Console.WriteLine($"Skipped password‑protected workbook: {Path.GetFileName(filePath)}");
                    }
                    else
                    {
                        Console.WriteLine($"CellsException processing '{Path.GetFileName(filePath)}': {ex.Message}");
                    }
                }
                catch (Exception ex)
                {
                    // Handle other unexpected errors gracefully
                    Console.WriteLine($"Error processing '{Path.GetFileName(filePath)}': {ex.Message}");
                }
            }

            // Ensure the output directory exists
            string? reportDir = Path.GetDirectoryName(reportPath);
            if (!string.IsNullOrEmpty(reportDir))
            {
                Directory.CreateDirectory(reportDir);
            }

            // Write all lines to the CSV file
            File.WriteAllLines(reportPath, reportLines);

            Console.WriteLine("Report generated at: " + reportPath);
        }
        catch (Exception ex)
        {
            // Top‑level safety net
            Console.WriteLine("Fatal error: " + ex.Message);
        }
    }
}