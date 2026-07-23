// Title: C# – Consolidate Built‑In Document Properties from Multiple Excel Workbooks with Aspose.Cells
// Description: Loads each Excel file in a supplied list, extracts every built‑in document property, and writes the source file name, property name, and value to a new worksheet. The resulting summary workbook is saved as BuiltInPropertiesSummary.xlsx and includes basic error handling for missing or unreadable files.
// Keywords: Aspose.Cells | C# | built‑in document properties | Excel metadata aggregation | property summary workbook | extract Excel file properties | batch document property report
// Common Searches: extract built‑in properties from multiple Excel files using Aspose.Cells | create a summary sheet of Excel file metadata in C# | combine document properties of several workbooks into one report | Aspose.Cells example for aggregating workbook properties | how to generate a property inventory for Excel files
// Developer Intent: Produce a single Excel file that lists each source workbook, its built‑in property names, and the corresponding values.
// Use Cases: Audit metadata across a batch of financial statements to verify author, company, and revision information. | Generate an inventory of document properties for compliance reporting on project deliverables. | Create a quick reference for property values before publishing or archiving a collection of workbooks.
// AI Prompts: Write C# code with Aspose.Cells that reads built‑in document properties from a list of Excel files and writes Source File, Property Name, and Value columns to a new workbook. | Extend the sample to also include custom document properties in the same summary sheet. | Add logging that records missing files and continues processing without stopping the application.

using System;
using System.IO;
using Aspose.Cells;

namespace SummaryReportExample
{
    // Loads each Excel file in a supplied list, extracts every built‑in document property, and writes the source file name, property name, and value to a new worksheet. The resulting summary workbook is saved as BuiltInPropertiesSummary.xlsx and includes basic error handling for missing or unreadable files.
    class Program
    {
        static void Main()
        {
            // List of workbook file paths to aggregate built‑in properties from
            string[] sourceFiles = new string[]
            {
                "Workbook1.xlsx",
                "Workbook2.xlsx",
                "Workbook3.xlsx"
                // Add more file paths as needed
            };

            // Create a new workbook that will hold the summary report
            Workbook summaryWorkbook = new Workbook();
            Worksheet sheet = summaryWorkbook.Worksheets[0];

            // Write header row
            int currentRow = 0;
            sheet.Cells[currentRow, 0].PutValue("Source File");
            sheet.Cells[currentRow, 1].PutValue("Property Name");
            sheet.Cells[currentRow, 2].PutValue("Value");

            foreach (string filePath in sourceFiles)
            {
                try
                {
                    // Verify the source file exists
                    if (!File.Exists(filePath))
                    {
                        Console.WriteLine($"Warning: File not found – {filePath}. Skipping.");
                        continue;
                    }

                    // Load the source workbook
                    Workbook sourceWorkbook = new Workbook(filePath);

                    // Iterate through all built‑in document properties
                    foreach (var prop in sourceWorkbook.BuiltInDocumentProperties)
                    {
                        currentRow++;
                        sheet.Cells[currentRow, 0].PutValue(filePath);
                        sheet.Cells[currentRow, 1].PutValue(prop.Name);
                        sheet.Cells[currentRow, 2].PutValue(prop.Value?.ToString() ?? string.Empty);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
                }
            }

            try
            {
                // Save the consolidated summary workbook
                summaryWorkbook.Save("BuiltInPropertiesSummary.xlsx", SaveFormat.Xlsx);
                Console.WriteLine("Summary workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error saving summary workbook: {ex.Message}");
            }
        }
    }
}
