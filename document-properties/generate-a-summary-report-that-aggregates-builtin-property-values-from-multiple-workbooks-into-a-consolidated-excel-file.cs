// Title: Consolidate Built‑In Document Properties from Multiple Excel Workbooks into a Summary Sheet with Aspose.Cells (C#)
// Description: A C# example that loads several Excel files, extracts each workbook's built‑in document properties, and writes the workbook name, property name, and value to a new worksheet. The script handles missing files, logs errors, creates the output folder if needed, and saves a single summary workbook.
// Keywords: Aspose.Cells | C# | built-in document properties | Excel metadata extraction | aggregate workbook properties | summary report Excel | batch process Excel files | extract Excel metadata .NET | consolidate Excel properties | Aspose.Cells example
// Common Searches: how to extract built‑in properties from multiple Excel files using Aspose.Cells | C# code to create a summary workbook of Excel metadata | combine document properties of several workbooks into one sheet | Aspose.Cells aggregate workbook properties example | generate audit report of Excel file metadata in .NET
// Developer Intent: Collect the built‑in document properties from a set of Excel workbooks and write them into a single consolidated report workbook.
// Use Cases: Audit a collection of financial spreadsheets by listing author, creation date, and other metadata in one file. | Prepare a compliance checklist that shows key property values across all departmental workbooks. | Provide a quick overview of Excel file metadata before migrating a batch of spreadsheets to a new platform.
// AI Prompts: Generate C# code with Aspose.Cells that reads built‑in document properties from a list of Excel files and writes them to a summary sheet with columns Workbook, Property, Value. | Explain how to add robust error handling for missing files and load failures when aggregating Excel metadata with Aspose.Cells. | Show how to format the summary worksheet (bold header, auto‑fit columns, freeze top row) after populating the aggregated properties.

using System;
using System.IO;
using Aspose.Cells;

// A C# example that loads several Excel files, extracts each workbook's built‑in document properties, and writes the workbook name, property name, and value to a new worksheet. The script handles missing files, logs errors, creates the output folder if needed, and saves a single summary workbook.
class SummaryReportGenerator
{
    static void Main()
    {
        try
        {
            // Paths of the workbooks to aggregate built‑in properties from
            string[] sourceFiles = new string[]
            {
                @"C:\Data\Workbook1.xlsx",
                @"C:\Data\Workbook2.xlsx",
                @"C:\Data\Workbook3.xlsx"
            };

            // Create a new workbook that will hold the summary report
            using (Workbook summaryWorkbook = new Workbook())
            {
                Worksheet sheet = summaryWorkbook.Worksheets[0];

                // Write header row
                sheet.Cells[0, 0].PutValue("Workbook");
                sheet.Cells[0, 1].PutValue("Property");
                sheet.Cells[0, 2].PutValue("Value");

                int currentRow = 1; // start after header (zero‑based index)

                foreach (string filePath in sourceFiles)
                {
                    // Verify source file exists to avoid FileNotFoundException
                    if (!File.Exists(filePath))
                    {
                        Console.WriteLine($"Source file not found: {filePath}");
                        continue;
                    }

                    try
                    {
                        // Load each source workbook inside a using block for proper disposal
                        using (Workbook srcWorkbook = new Workbook(filePath))
                        {
                            // Get a friendly name for the workbook (file name only)
                            string workbookName = Path.GetFileName(filePath);

                            // Iterate through all built‑in document properties
                            foreach (var prop in srcWorkbook.BuiltInDocumentProperties)
                            {
                                // Write workbook name, property name and its value into the summary sheet
                                sheet.Cells[currentRow, 0].PutValue(workbookName);
                                sheet.Cells[currentRow, 1].PutValue(prop.Name);
                                sheet.Cells[currentRow, 2].PutValue(prop.Value?.ToString() ?? string.Empty);
                                currentRow++;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
                    }
                }

                // Ensure the output directory exists
                string outputPath = @"C:\Data\SummaryReport.xlsx";
                string? outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the consolidated summary workbook
                summaryWorkbook.Save(outputPath);
            }

            Console.WriteLine("Summary report generated successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
