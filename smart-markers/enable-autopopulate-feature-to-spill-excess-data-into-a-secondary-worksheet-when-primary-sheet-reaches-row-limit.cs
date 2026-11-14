// Title: How to automatically spill excess rows to a secondary worksheet when the primary sheet reaches Excel’s 1,048,576‑row limit using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# write data beyond Excel row limit to another sheet | C# split large dataset across multiple worksheets with Aspose.Cells | How to handle Excel max rows in Aspose.Cells by creating overflow worksheet | Aspose.Cells example for spilling rows to secondary worksheet after row limit | Write 2 million rows to Excel using Aspose.Cells and overflow to second sheet
// Tags: spill rows to secondary worksheet Aspose.Cells | handle Excel max row limit C# | auto overflow rows Aspose.Cells | multiple worksheets data export Aspose.Cells | large dataset Excel generation Aspose.Cells

using Aspose.Cells;
using System;
using System.Collections.Generic;

// The sample defines Excel's maximum row count, generates two million rows, creates a workbook with a primary and a secondary worksheet, writes headers, fills the primary sheet until the 1,048,576‑row limit is reached, then continues writing remaining rows to the secondary sheet, and saves the file as Spillover.xlsx.
class Program
{
    static void Main()
    {
        // Maximum rows per worksheet in Excel (1,048,576)
        const int MaxRows = 1048576;

        // Example data source with more rows than the limit
        List<string[]> data = GenerateData(2000000); // 2 million rows

        // Create a new workbook (primary sheet is created by default)
        Workbook workbook = new Workbook();

        // Rename the default sheet to "Primary"
        Worksheet primarySheet = workbook.Worksheets[0];
        primarySheet.Name = "Primary";

        // Add a secondary worksheet to receive overflow rows
        Worksheet secondarySheet = workbook.Worksheets.Add("Secondary");

        // Optional: write column headers on both sheets
        primarySheet.Cells[0, 0].PutValue("Column1");
        primarySheet.Cells[0, 1].PutValue("Column2");
        secondarySheet.Cells[0, 0].PutValue("Column1");
        secondarySheet.Cells[0, 1].PutValue("Column2");

        // Start writing data after the header row
        int primaryRow = 1;
        int secondaryRow = 1;

        foreach (var row in data)
        {
            if (primaryRow < MaxRows)
            {
                // Write to primary sheet while within row limit
                primarySheet.Cells[primaryRow, 0].PutValue(row[0]);
                primarySheet.Cells[primaryRow, 1].PutValue(row[1]);
                primaryRow++;
            }
            else
            {
                // Spill excess rows into secondary sheet
                secondarySheet.Cells[secondaryRow, 0].PutValue(row[0]);
                secondarySheet.Cells[secondaryRow, 1].PutValue(row[1]);
                secondaryRow++;
            }
        }

        // Save the workbook to a file
        workbook.Save("Spillover.xlsx");
    }

    // Generates dummy data for demonstration purposes
    static List<string[]> GenerateData(int rowCount)
    {
        var list = new List<string[]>(rowCount);
        for (int i = 0; i < rowCount; i++)
        {
            list.Add(new string[] { $"Item {i + 1}", $"Value {i + 1}" });
        }
        return list;
    }
}
