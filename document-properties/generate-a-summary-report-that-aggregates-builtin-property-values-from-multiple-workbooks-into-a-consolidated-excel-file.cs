using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace ConsolidatedSummaryReport
{
    class Program
    {
        static void Main()
        {
            try
            {
                // List of source workbook file paths
                List<string> sourceFiles = new List<string>
                {
                    @"C:\Data\Workbook1.xlsx",
                    @"C:\Data\Workbook2.xlsx",
                    @"C:\Data\Workbook3.xlsx"
                    // Add more paths as needed
                };

                // Create a new workbook that will hold the consolidated summary
                Workbook summaryWorkbook = new Workbook();

                // Use the first worksheet as the summary sheet
                Worksheet summarySheet = summaryWorkbook.Worksheets[0];

                // Write header row
                summarySheet.Cells[0, 0].PutValue("Workbook");
                summarySheet.Cells[0, 1].PutValue("Property");
                summarySheet.Cells[0, 2].PutValue("Value");

                int currentRow = 1; // start after header

                foreach (string filePath in sourceFiles)
                {
                    // Skip missing files to avoid FileNotFoundException
                    if (!File.Exists(filePath))
                    {
                        Console.WriteLine($"File not found: {filePath}");
                        continue;
                    }

                    try
                    {
                        // Load each source workbook
                        Workbook sourceWorkbook = new Workbook(filePath);

                        // Get the file name for display
                        string workbookName = Path.GetFileName(filePath);

                        // Iterate through built‑in document properties
                        foreach (var prop in sourceWorkbook.BuiltInDocumentProperties)
                        {
                            // Write workbook name, property name, and property value to the summary sheet
                            summarySheet.Cells[currentRow, 0].PutValue(workbookName);
                            summarySheet.Cells[currentRow, 1].PutValue(prop.Name);
                            summarySheet.Cells[currentRow, 2].PutValue(prop.Value?.ToString() ?? string.Empty);
                            currentRow++;
                        }
                    }
                    catch (Exception ex)
                    {
                        // Log errors for individual workbooks but continue processing others
                        Console.WriteLine($"Error processing '{filePath}': {ex.Message}");
                    }
                }

                // Ensure the output directory exists
                string outputPath = @"C:\Data\ConsolidatedSummary.xlsx";
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the consolidated summary workbook
                summaryWorkbook.Save(outputPath);
                Console.WriteLine($"Consolidated summary saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}