// Title: Batch create line sparklines for rows 1‑20 in Excel using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that uses Aspose.Cells to add a SparklineGroup of type Line and inserts a sparkline for each of the first 20 rows, pulling data from columns A‑D and placing the sparkline in column E. | Write a script that populates rows 1‑20 with sample numeric values, then programmatically creates sparklines for each row using the ranges A1:D1, A2:D2, …, and saves the workbook as an .xlsx file. | Provide a step‑by‑step example showing how to loop through worksheet rows and call SparklineGroup.Sparklines.Add for batch sparkline creation with Aspose.Cells.
// Common Searches: Aspose.Cells C# how to add a sparkline for each row in a worksheet | batch generate line sparklines for rows 1 to 20 using .NET | programmatically create sparklines from row data range A:D in Excel with Aspose | save Excel file with sparklines using Aspose.Cells for .NET | C# example of SparklineGroup.Add for multiple rows
// Tags: batch add line sparklines Aspose.Cells C# | populate worksheet rows for sparkline data Aspose.Cells | create SparklineGroup programmatically .NET | add sparklines to column E Aspose.Cells | save workbook with sparklines .xlsx Aspose

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a new workbook, fills rows 1‑20 (columns A‑D) with sample numeric data, adds a line SparklineGroup, then loops through each row to add a sparkline that references the row's A‑D range and places it in column E. Finally, the workbook is saved as BatchSparklines.xlsx.
class BatchSparklineDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // ------------------------------------------------------------
        // Populate sample data for rows 1‑20, columns A‑D (4 columns)
        // ------------------------------------------------------------
        for (int row = 0; row < 20; row++)          // zero‑based index
        {
            for (int col = 0; col < 4; col++)       // columns A‑D
            {
                // Example data: (row index + 1) * (col index + 1)
                sheet.Cells[row, col].PutValue((row + 1) * (col + 1));
            }
        }

        // ------------------------------------------------------------
        // Add a SparklineGroup of type Line (you can change the type)
        // ------------------------------------------------------------
        int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line);
        SparklineGroup group = sheet.SparklineGroups[groupIndex];

        // ------------------------------------------------------------
        // Batch create sparklines for rows 1‑20.
        // Each sparkline uses the data range of its own row (A‑D)
        // and is placed in column E (zero‑based index 4) of the same row.
        // ------------------------------------------------------------
        for (int row = 0; row < 20; row++)
        {
            // Build the data range string for the current row, e.g., "A1:D1"
            string dataRange = $"A{row + 1}:D{row + 1}";

            // Add the sparkline: dataRange, row index, column index (E = 4)
            group.Sparklines.Add(dataRange, row, 4);
        }

        // ------------------------------------------------------------
        // Save the workbook with the created sparklines
        // ------------------------------------------------------------
        workbook.Save("BatchSparklines.xlsx");
    }
}
