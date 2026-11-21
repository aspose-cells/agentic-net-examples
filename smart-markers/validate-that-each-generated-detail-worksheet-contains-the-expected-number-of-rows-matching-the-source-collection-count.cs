// Title: C# code to validate that Aspose.Cells generated Detail worksheets contain the expected number of rows matching their source collections
// Common Searches: how to verify Excel worksheet row count matches a list size using Aspose.Cells C# | Aspose.Cells validate detail sheet rows against source collection in .NET | C# compare MaxDataRow value with collection count for generated Excel sheets | check if Aspose.Cells detail worksheets contain expected number of rows | using a dictionary of source data to validate Excel detail worksheets in C#
// Tags: Aspose.Cells row count validation | C# Excel worksheet row verification | MaxDataRow collection count comparison | detail worksheet verification Aspose.Cells | smart markers generated sheet validation

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, iterates over worksheets whose names start with "Detail", retrieves the corresponding source collection from a predefined dictionary, and compares the actual data rows (using MaxDataRow + 1) to the collection's item count, reporting any mismatches.
class DetailWorksheetValidator
{
    // Mapping of worksheet names to their corresponding source collections.
    // Populate this dictionary with the actual data used to generate each detail sheet.
    private static readonly Dictionary<string, IList<object>> SourceDataMap = new Dictionary<string, IList<object>>
    {
        // Example entries:
        // { "DetailSales", salesDataList },
        // { "DetailInventory", inventoryDataList }
    };

    static void Main()
    {
        const string filePath = "GeneratedReport.xlsx";
        Workbook workbook = null;

        try
        {
            // Verify that the input file exists before attempting to load it.
            if (File.Exists(filePath))
            {
                workbook = new Workbook(filePath);
            }
            else
            {
                Console.WriteLine($"File '{filePath}' not found. Creating an empty workbook for validation.");
                workbook = new Workbook(); // creates a new workbook with a default sheet
            }

            // Iterate through all worksheets in the workbook.
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Process only worksheets that are considered "detail" sheets.
                if (!sheet.Name.StartsWith("Detail", StringComparison.OrdinalIgnoreCase))
                    continue;

                // Retrieve the source collection associated with this worksheet.
                if (!SourceDataMap.TryGetValue(sheet.Name, out IList<object> sourceCollection))
                {
                    Console.WriteLine($"Warning: No source collection found for worksheet '{sheet.Name}'. Skipping validation.");
                    continue;
                }

                // Determine the number of rows that contain data in the worksheet.
                // MaxDataRow returns the zero‑based index of the last row that has data.
                int actualRowCount = sheet.Cells.MaxDataRow + 1; // Convert to 1‑based count.

                // Expected row count is the number of items in the source collection.
                int expectedRowCount = sourceCollection.Count;

                // Validate the row count.
                if (actualRowCount != expectedRowCount)
                {
                    Console.WriteLine($"Validation failed for worksheet '{sheet.Name}': Expected {expectedRowCount} rows, but found {actualRowCount} rows.");
                }
                else
                {
                    Console.WriteLine($"Worksheet '{sheet.Name}' passed validation (Rows: {actualRowCount}).");
                }
            }

            // Optionally, save the workbook after validation (if any modifications were made).
            // workbook.Save("ValidatedReport.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred during validation: {ex.Message}");
        }
    }
}
