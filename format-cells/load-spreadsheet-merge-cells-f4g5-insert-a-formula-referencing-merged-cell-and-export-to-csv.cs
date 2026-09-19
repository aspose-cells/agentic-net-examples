// Title: Merge cells F4:G5, assign a formula to the merged range, and export the worksheet as CSV using Aspose.Cells for .NET
// AI Prompts: Use Aspose.Cells in C# to load an Excel file, merge the range F4:G5, set cell A1 formula to reference the merged top‑left cell, and save the result as a CSV file. | Write C# code that demonstrates merging cells, inserting a formula that points to the merged cell, and converting the worksheet to CSV with Aspose.Cells.
// Common Searches: Aspose.Cells C# merge a specific range and export to CSV | How to reference a merged cell in a formula with Aspose.Cells .NET | Save Excel worksheet as CSV after merging cells using Aspose.Cells | C# example for merging cells F4:G5 and setting formula =F4 with Aspose.Cells | Export merged cell data to CSV with Aspose.Cells for .NET
// Tags: merge cell range F4:G5 Aspose.Cells | set formula referencing merged cell Aspose.Cells | export worksheet to CSV Aspose.Cells | Aspose.Cells C# cell merging and formula | CSV conversion after cell merge Aspose.Cells

using System;
using Aspose.Cells;

// Loads input.xlsx, merges cells F4:G5, assigns A1 a formula that points to the merged top‑left cell, and saves the worksheet as output.csv in CSV format using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Load the existing spreadsheet
        Workbook workbook = new Workbook("input.xlsx");

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Merge cells F4:G5 (zero‑based indices: row 3, column 5, 2 rows, 2 columns)
        sheet.Cells.Merge(3, 5, 2, 2);

        // Insert a formula that references the merged cell (top‑left cell of the merged range)
        sheet.Cells["A1"].Formula = "=F4";

        // Export the worksheet to CSV format
        workbook.Save("output.csv", SaveFormat.CSV);
    }
}
