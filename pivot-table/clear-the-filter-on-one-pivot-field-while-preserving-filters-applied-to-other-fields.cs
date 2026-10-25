// Title: Clearing a filter on a specific pivot page field while preserving other filters in an Excel workbook using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# clear filter on one pivot page field only | remove filter from a specific pivot field without affecting other filters Aspose.Cells | how to reset a single page field filter in a pivot table using Aspose.Cells .NET | C# code to clear pivot table page field filter while keeping other filters | Aspose.Cells example to clear pivot table filter on selected field
// Tags: Aspose.Cells clear pivot page field filter | Aspose.Cells preserve other pivot filters | C# pivot table filter manipulation Aspose.Cells | Excel workbook pivot field filter reset .NET | Aspose.Cells pivot table page fields API

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // The example loads an Excel workbook, accesses a named pivot table, validates a page‑field index, clears the filter on that specific pivot field by disabling its auto‑show and auto‑sort settings, and saves the workbook, leaving filters on all other pivot fields unchanged.
    class Program
    {
        static void Main(string[] args)
        {
            // Define file paths and parameters
            string workbookPath = "input.xlsx";          // TODO: set actual input workbook path
            string pivotTableName = "PivotTable1";       // TODO: set actual pivot table name
            int fieldIndex = 0;                          // Index of the page field to clear filter (0‑based)
            string outputPath = "output.xlsx";           // TODO: set desired output path

            try
            {
                // Verify that the input workbook exists
                if (!File.Exists(workbookPath))
                {
                    Console.WriteLine($"Error: Workbook file not found at '{workbookPath}'.");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(workbookPath);

                // Access the first worksheet (adjust if needed)
                Worksheet worksheet = workbook.Worksheets[0];

                // Retrieve the pivot table by name
                PivotTable pivotTable = worksheet.PivotTables[pivotTableName];
                if (pivotTable == null)
                {
                    Console.WriteLine($"Error: Pivot table '{pivotTableName}' not found in worksheet '{worksheet.Name}'.");
                    return;
                }

                // Ensure the field index is valid
                if (fieldIndex < 0 || fieldIndex >= pivotTable.PageFields.Count)
                {
                    Console.WriteLine($"Error: Field index {fieldIndex} is out of range. The pivot table has {pivotTable.PageFields.Count} page fields.");
                    return;
                }

                // Clear the filter on the selected page field
                PivotField targetField = pivotTable.PageFields[fieldIndex];
                targetField.IsAutoShow = false;
                targetField.IsAutoSort = false;
                // Note: Aspose.Cells does not expose an IsAutoSubtotal property for PivotField; omitted.

                // Ensure output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An unexpected error occurred:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
