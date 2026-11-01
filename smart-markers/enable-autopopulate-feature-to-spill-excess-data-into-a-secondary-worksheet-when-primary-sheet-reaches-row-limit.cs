// Title: How to automatically spill rows to an overflow worksheet when a primary sheet reaches a row limit using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that fills a primary worksheet up to a configurable row count and then continues inserting rows into a second worksheet. | Create a method that accepts a list of string arrays and a maximum rows‑per‑sheet value, and generates an Excel file where excess rows are written to an overflow sheet using Aspose.Cells. | Generate a sample workbook that demonstrates splitting 2500 rows between "PrimaryData" and "OverflowData" worksheets based on a 1000‑row threshold.
// Common Searches: Aspose.Cells C# write data to multiple worksheets after reaching row limit | How to split Excel rows into an overflow sheet with Aspose.Cells .NET | C# Aspose.Cells auto move rows to secondary worksheet when first sheet is full | limit rows per worksheet Aspose.Cells example | auto populate overflow worksheet when primary sheet exceeds row count
// Tags: spill rows to overflow worksheet Aspose.Cells | row limit per sheet C# Aspose.Cells | auto populate secondary worksheet Aspose.Cells | split data across worksheets Aspose.Cells .NET | write large dataset to multiple Excel sheets Aspose.Cells

using System;
using System.Collections.Generic;
using Aspose.Cells;

// The example creates a new workbook, adds a primary sheet named "PrimaryData" and an overflow sheet named "OverflowData", generates sample data, writes rows to the primary sheet up to a defined maximum (e.g., 1000 rows), then automatically continues writing remaining rows to the overflow sheet, and finally saves the workbook as AutoPopulateSpillResult.xlsx.
class AutoPopulateSpillExample
{
    static void Main()
    {
        // Create a new workbook (lifecycle rule: create)
        Workbook workbook = new Workbook();

        // Access the primary worksheet (default sheet)
        Worksheet primarySheet = workbook.Worksheets[0];
        primarySheet.Name = "PrimaryData";

        // Add a secondary worksheet for overflow data
        Worksheet overflowSheet = workbook.Worksheets.Add("OverflowData");

        // Define the maximum number of rows allowed in the primary sheet
        // (e.g., 1000 rows for demonstration; adjust as needed)
        const int maxPrimaryRows = 1000;

        // Sample data to populate (replace with actual data source)
        List<string[]> data = GenerateSampleData(2500); // 2500 rows of sample data

        // Populate data, spilling excess rows to the overflow sheet
        int primaryRowIndex = 0;   // zero‑based index for primary sheet
        int overflowRowIndex = 0;  // zero‑based index for overflow sheet

        foreach (var row in data)
        {
            // Determine target worksheet based on current row index
            Worksheet targetSheet;
            int targetRowIndex;

            if (primaryRowIndex < maxPrimaryRows)
            {
                targetSheet = primarySheet;
                targetRowIndex = primaryRowIndex;
                primaryRowIndex++;
            }
            else
            {
                targetSheet = overflowSheet;
                targetRowIndex = overflowRowIndex;
                overflowRowIndex++;
            }

            // Write each column value into the target worksheet
            for (int col = 0; col < row.Length; col++)
            {
                // Cells are addressed by zero‑based indices
                targetSheet.Cells[targetRowIndex, col].PutValue(row[col]);
            }
        }

        // Save the workbook to a file (lifecycle rule: save)
        workbook.Save("AutoPopulateSpillResult.xlsx");
    }

    // Helper method to generate sample data (rows x columns)
    static List<string[]> GenerateSampleData(int rowCount)
    {
        var list = new List<string[]>();
        for (int i = 1; i <= rowCount; i++)
        {
            // Example: three columns per row
            list.Add(new string[]
            {
                $"Row{i}",
                $"ValueA{i}",
                $"ValueB{i}"
            });
        }
        return list;
    }
}
