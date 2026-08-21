// Title: C# Batch Processing of 100 Excel Workbooks with Aspose.Cells – Add Metadata & Clean Unused Styles
// Description: A ready‑to‑run C# console example that loops through 100 Excel files, loads each workbook with Aspose.Cells, optionally adds a custom document property and sets the Author built‑in property, removes any unused cell styles to shrink file size, and saves the result to a separate output folder. The code includes checks for missing files and robust error handling, making it ideal for automation scripts or GitHub‑hosted utilities.
// Keywords: Aspose.Cells batch processing | C# Excel workbook loop | add custom document property Aspose | remove unused styles Excel | bulk Excel metadata update | automate Excel file cleanup .NET | GitHub Aspose.Cells example | Excel file size reduction
// Common Searches: how to process multiple Excel files with Aspose.Cells in C# | add custom document properties to many workbooks programmatically | remove unused styles from Excel workbooks in bulk | set author property for a batch of Excel files using .NET | sample code for batch Excel processing Aspose.Cells GitHub
// Developer Intent: Automate the modification of a large set of Excel workbooks—add optional metadata and eliminate unused styles—while handling missing files and runtime errors.
// Use Cases: Prepare 100 workbooks for archiving by inserting a custom property that records the batch job identifier. | Standardize the Author property across all files before publishing or compliance checks. | Reduce workbook size and improve performance by stripping unused cell styles during bulk conversion or migration.
// AI Prompts: Generate C# code using Aspose.Cells to iterate over a folder of Excel files, add custom and built‑in document properties, remove unused styles, and save each file to an output directory. | Suggest best‑practice error‑handling patterns for large‑scale workbook processing with Aspose.Cells in .NET. | Explain how to verify that RemoveUnusedStyles() successfully eliminated all redundant styles after processing.

using System;
using System.IO;
using Aspose.Cells;

namespace BatchWorkbookProcessor
{
    // A ready‑to‑run C# console example that loops through 100 Excel files, loads each workbook with Aspose.Cells, optionally adds a custom document property and sets the Author built‑in property, removes any unused cell styles to shrink file size, and saves the result to a separate output folder. The code includes checks for missing files and robust error handling, making it ideal for automation scripts or GitHub‑hosted utilities.
    class Program
    {
        static void Main()
        {
            // Folder paths – adjust as needed
            string inputFolder = @"C:\Workbooks\Input";
            string outputFolder = @"C:\Workbooks\Output";

            // Ensure the output directory exists
            Directory.CreateDirectory(outputFolder);

            // Process 100 workbooks
            for (int i = 1; i <= 100; i++)
            {
                string inputFile = Path.Combine(inputFolder, $"Workbook{i}.xlsx");
                string outputFile = Path.Combine(outputFolder, $"Workbook{i}_processed.xlsx");

                // Skip missing input files
                if (!File.Exists(inputFile))
                {
                    Console.WriteLine($"Input file not found: {inputFile}");
                    continue;
                }

                try
                {
                    // Load the workbook
                    Workbook wb = new Workbook(inputFile);

                    // Add custom metadata (optional)
                    wb.CustomDocumentProperties.Add($"ProcessedBy_{i}", "BatchJob");

                    // Add built‑in metadata (optional)
                    wb.BuiltInDocumentProperties["Author"].Value = "AutomationEngine";

                    // Remove unused styles
                    wb.RemoveUnusedStyles();

                    // Save the modified workbook
                    wb.Save(outputFile);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing workbook {i}: {ex.Message}");
                }
            }

            Console.WriteLine("Batch processing of workbooks completed.");
        }
    }
}
