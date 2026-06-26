using System;
using System.IO;
using Aspose.Cells;

class ConsolidateBuiltInProperties
{
    static void Main()
    {
        try
        {
            // Paths of source workbooks to be aggregated
            string[] sourceFiles = new string[]
            {
                @"C:\Data\Workbook1.xlsx",
                @"C:\Data\Workbook2.xlsx",
                @"C:\Data\Workbook3.xlsx"
            };

            // Create a new workbook that will hold the summary report
            Workbook summaryWorkbook = new Workbook(); // lifecycle: create
            Worksheet summarySheet = summaryWorkbook.Worksheets[0];
            summarySheet.Name = "Summary";

            // Write header row
            summarySheet.Cells["A1"].PutValue("Workbook");
            summarySheet.Cells["B1"].PutValue("Property");
            summarySheet.Cells["C1"].PutValue("Value");

            int currentRow = 1; // zero‑based index; row 1 is the second row (after header)

            foreach (string filePath in sourceFiles)
            {
                // Verify source file exists
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"Warning: Source file not found and will be skipped: {filePath}");
                    continue;
                }

                try
                {
                    // Load each source workbook (lifecycle: load)
                    using (Workbook sourceWorkbook = new Workbook(filePath))
                    {
                        string workbookName = Path.GetFileName(filePath);

                        // Iterate through all built‑in document properties
                        foreach (var prop in sourceWorkbook.BuiltInDocumentProperties)
                        {
                            // Write workbook name, property name and its value into the summary sheet
                            summarySheet.Cells[currentRow, 0].PutValue(workbookName);               // Column A
                            summarySheet.Cells[currentRow, 1].PutValue(prop.Name);                 // Column B
                            summarySheet.Cells[currentRow, 2].PutValue(prop.Value?.ToString() ?? string.Empty); // Column C
                            currentRow++;
                        }
                    } // sourceWorkbook disposed here
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
                }
            }

            // Ensure output directory exists
            string outputPath = @"C:\Data\BuiltInPropertiesSummary.xlsx";
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            try
            {
                // Save the consolidated summary workbook (lifecycle: save)
                summaryWorkbook.Save(outputPath);
                Console.WriteLine($"Summary workbook saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save summary workbook: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}