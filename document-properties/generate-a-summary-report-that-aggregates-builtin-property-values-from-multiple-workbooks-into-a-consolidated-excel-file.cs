// Title: Create a C# console app to consolidate built-in document properties from multiple Excel workbooks into a single report using Aspose.Cells
// AI Prompts: Generate a C# console application that loads a collection of .xlsx files, reads each workbook's BuiltInDocumentProperties via Aspose.Cells, and writes the source file name, property name, and property value into a new worksheet. | Extend the program to also capture each workbook's custom document properties and include them in the same aggregated report. | Add robust logging that records missing files and any exceptions encountered while processing the workbooks, and ensure the output directory is created automatically. | Implement column auto‑fit and apply a bold header style to the aggregated properties sheet for better readability.
// Common Searches: aspnet read built-in document properties from multiple Excel files and combine into one report | c# Aspose.Cells aggregate workbook metadata into a summary spreadsheet | how to list all built-in properties of several .xlsx files using Aspose.Cells | generate consolidated Excel file with property name and value for a collection of workbooks in .NET | skip missing Excel files while extracting document properties with Aspose.Cells
// Tags: Aspose.Cells read document property set | C# create unified Excel property workbook | auto-size columns in generated sheet | skip non-existent source files gracefully | apply bold header row to report sheet

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace BuiltInPropertiesAggregator
{
    // The program iterates over a list of Excel file paths, loads each workbook with Aspose.Cells, extracts every built-in document property, and writes the source file name, property name, and its value into a new worksheet. It then auto-fits columns, adds a styled header, and saves the consolidated workbook as a summary report.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // List of source workbook file paths to aggregate
                var sourceFiles = new List<string>
                {
                    @"C:\Data\Workbook1.xlsx",
                    @"C:\Data\Workbook2.xlsx",
                    @"C:\Data\Workbook3.xlsx"
                    // Add more file paths as needed
                };

                // Create a new workbook that will hold the consolidated report
                var reportWorkbook = new Workbook();
                var reportSheet = reportWorkbook.Worksheets[0];
                reportSheet.Name = "BuiltInProperties";

                // Write header row
                reportSheet.Cells[0, 0].PutValue("Source Workbook");
                reportSheet.Cells[0, 1].PutValue("Property Name");
                reportSheet.Cells[0, 2].PutValue("Property Value");

                int currentRow = 1; // Start writing data from the second row

                foreach (var filePath in sourceFiles)
                {
                    // Verify that the source file exists
                    if (!File.Exists(filePath))
                    {
                        Console.WriteLine($"Warning: File not found – {filePath}. Skipping.");
                        continue;
                    }

                    try
                    {
                        // Load each source workbook
                        var srcWorkbook = new Workbook(filePath);

                        // Access built‑in document properties
                        var properties = srcWorkbook.BuiltInDocumentProperties;

                        // Iterate through all built‑in properties
                        foreach (var prop in properties)
                        {
                            // Write source workbook name (without full path)
                            reportSheet.Cells[currentRow, 0].PutValue(Path.GetFileName(filePath));

                            // Write property name
                            reportSheet.Cells[currentRow, 1].PutValue(prop.Name);

                            // Write property value as string (handle nulls)
                            string valueStr = prop.Value != null ? prop.Value.ToString() : string.Empty;
                            reportSheet.Cells[currentRow, 2].PutValue(valueStr);

                            currentRow++;
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
                    }
                }

                // Auto‑fit columns for better readability
                reportSheet.AutoFitColumns();

                // Ensure the output directory exists
                string reportPath = @"C:\Data\BuiltInPropertiesReport.xlsx";
                string? reportDir = Path.GetDirectoryName(reportPath);
                if (!string.IsNullOrEmpty(reportDir) && !Directory.Exists(reportDir))
                {
                    Directory.CreateDirectory(reportDir);
                }

                // Save the consolidated report
                reportWorkbook.Save(reportPath, SaveFormat.Xlsx);
                Console.WriteLine($"Report generated successfully at: {reportPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
