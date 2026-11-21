// Title: Enforce a maximum row height after AutoFitRows with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# limit row height after AutoFitRows | Check and clamp Excel row height to 50 points using Aspose.Cells .NET | Validate row heights after auto‑fit in a workbook with Aspose.Cells | Set maximum row height when auto‑fitting rows in Aspose.Cells for .NET | Programmatically adjust rows that exceed height limit after AutoFitRows in C#
// Tags: auto-fit rows maximum height Aspose.Cells | row height clamping after AutoFitRows C# | validate row height limit .NET workbook | programmatic row height adjustment Excel Aspose.Cells | iterate worksheet rows enforce height constraint

using System;
using Aspose.Cells;

// The example loads an Excel workbook, auto‑fits all rows in each worksheet, iterates through each row, and resets any row height exceeding 50 points to that limit, logging the changes before saving the updated file.
class Program
{
    static void Main()
    {
        // Load the workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Maximum allowed row height in points
        const double maxRowHeight = 50.0;

        // Process each worksheet
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Auto‑fit all rows in the current sheet
            sheet.AutoFitRows();

            // Check each row's height after auto‑fit
            foreach (Row row in sheet.Cells.Rows)
            {
                // Skip rows that were never accessed (null)
                if (row == null) continue;

                // If the row height exceeds the limit, adjust it
                if (row.Height > maxRowHeight)
                {
                    // Clamp the height to the maximum allowed value
                    row.Height = maxRowHeight;

                    // Optional: output information for debugging/validation
                    Console.WriteLine($"Row {row.Index} in sheet \"{sheet.Name}\" exceeded max height. Height set to {maxRowHeight}.");
                }
            }
        }

        // Save the modified workbook (replace with your desired output path)
        workbook.Save("output.xlsx");
    }
}
