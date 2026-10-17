// Title: Create a consolidated Excel report of built-in document properties from multiple workbooks using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# read built-in document properties from several Excel files and write them to a summary sheet | How to generate a workbook metadata summary across multiple .xlsx files with Aspose.Cells | C# code to aggregate Excel built-in properties into one report workbook | Combine document properties of multiple spreadsheets into a single Excel file using Aspose.Cells | Export built-in properties of many workbooks to a consolidated .xlsx with Aspose.Cells .NET
// Tags: Aspose.Cells read built-in document properties | aggregate Excel workbook metadata C# | consolidate built-in properties into summary worksheet | export workbook properties to new Excel file | auto-fit columns Aspose.Cells summary report

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace BuiltInPropertiesAggregator
{
    // The example loads a list of Excel workbooks, extracts each workbook's BuiltInDocumentProperties via Aspose.Cells, and writes the workbook name, property name, and property value into a new "Summary" worksheet. It auto-fits columns and saves the consolidated report, handling missing files and save errors.
    class Program
    {
        static void Main(string[] args)
        {
            // List of source workbook file paths to aggregate.
            // Replace with actual paths or populate dynamically.
            List<string> sourceFiles = new List<string>
            {
                @"C:\Data\Workbook1.xlsx",
                @"C:\Data\Workbook2.xlsx",
                @"C:\Data\Workbook3.xlsx"
            };

            // Create a new workbook that will hold the consolidated report.
            Workbook summaryWorkbook = new Workbook();

            // Add a worksheet named "Summary".
            Worksheet summarySheet = summaryWorkbook.Worksheets[0];
            summarySheet.Name = "Summary";

            // Write header row.
            Cells cells = summarySheet.Cells;
            cells["A1"].PutValue("Workbook");
            cells["B1"].PutValue("Property Name");
            cells["C1"].PutValue("Value");

            int currentRow = 1; // Zero‑based index; row 1 is the second row (after header).

            foreach (string filePath in sourceFiles)
            {
                try
                {
                    // Ensure the source file exists before attempting to load.
                    if (!File.Exists(filePath))
                    {
                        Console.WriteLine($"Source file not found: {filePath}");
                        continue;
                    }

                    // Load the source workbook.
                    Workbook srcWorkbook = new Workbook(filePath);

                    // Access built‑in document properties.
                    var props = srcWorkbook.BuiltInDocumentProperties;

                    // Iterate through all built‑in properties.
                    foreach (var prop in props)
                    {
                        // Write workbook name, property name, and property value to the summary sheet.
                        cells[currentRow, 0].PutValue(Path.GetFileName(filePath));
                        cells[currentRow, 1].PutValue(prop.Name);
                        cells[currentRow, 2].PutValue(prop.Value?.ToString() ?? string.Empty);
                        currentRow++;
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
                }
            }

            // Auto‑fit columns for better readability.
            summarySheet.AutoFitColumns();

            // Save the consolidated report.
            string outputPath = @"C:\Data\BuiltInPropertiesSummary.xlsx";
            try
            {
                // Ensure the output directory exists.
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                summaryWorkbook.Save(outputPath);
                Console.WriteLine($"Summary report generated at: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save summary workbook: {ex.Message}");
            }
        }
    }
}
