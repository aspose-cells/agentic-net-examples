// Title: C# code to aggregate built‑in document properties from multiple Excel workbooks into a single summary worksheet using Aspose.Cells
// Common Searches: Aspose.Cells C# read built‑in document properties from a list of Excel files | How to create a summary worksheet that lists workbook metadata with Aspose.Cells .NET | C# aggregate Excel file properties into one report using Aspose.Cells | Generate a consolidated built‑in properties report for several XLSX workbooks in C# | Iterate over multiple workbooks and export document properties to a new Excel file with Aspose.Cells
// Tags: read built-in document properties Aspose.Cells | write summary worksheet Aspose.Cells | aggregate workbook metadata C# | consolidate Excel properties Aspose.Cells | auto-fit columns Aspose.Cells | load multiple workbooks Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace BuiltInPropertiesAggregator
{
    // The example loads a collection of Excel files, extracts each workbook's BuiltInDocumentProperties, and writes the workbook name, property name, and value to a new workbook's "Summary" sheet. It adds a header row, appends rows for every property, auto‑fits columns for readability, and saves the consolidated report as an XLSX file while handling missing files and processing errors.
    class Program
    {
        static void Main(string[] args)
        {
            // List of source workbook file paths to aggregate.
            // Replace these paths with the actual locations of your Excel files.
            List<string> sourceFiles = new List<string>
            {
                @"C:\Data\Workbook1.xlsx",
                @"C:\Data\Workbook2.xlsx",
                @"C:\Data\Workbook3.xlsx"
            };

            // Create a new workbook that will hold the consolidated report.
            Workbook summaryWorkbook = new Workbook();

            // Access the first (default) worksheet and rename it to "Summary".
            Worksheet summarySheet = summaryWorkbook.Worksheets[0];
            summarySheet.Name = "Summary";

            // Write header row.
            summarySheet.Cells[0, 0].PutValue("Workbook");
            summarySheet.Cells[0, 1].PutValue("Property Name");
            summarySheet.Cells[0, 2].PutValue("Value");

            int currentRow = 1; // Start writing data from the second row.

            // Iterate over each source workbook.
            foreach (string filePath in sourceFiles)
            {
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found: {filePath}");
                    continue;
                }

                try
                {
                    // Load the source workbook.
                    Workbook sourceWorkbook = new Workbook(filePath);

                    // Retrieve built‑in document properties.
                    var properties = sourceWorkbook.BuiltInDocumentProperties;

                    // Iterate through all built‑in properties.
                    foreach (var prop in properties)
                    {
                        // Write workbook name, property name, and property value to the summary sheet.
                        summarySheet.Cells[currentRow, 0].PutValue(Path.GetFileName(filePath));
                        summarySheet.Cells[currentRow, 1].PutValue(prop.Name);
                        summarySheet.Cells[currentRow, 2].PutValue(prop.Value?.ToString() ?? string.Empty);
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

            // Save the consolidated report to a new Excel file.
            string outputPath = @"C:\Data\BuiltInPropertiesSummary.xlsx";

            try
            {
                summaryWorkbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Summary report generated successfully at: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save summary workbook: {ex.Message}");
            }
        }
    }
}
