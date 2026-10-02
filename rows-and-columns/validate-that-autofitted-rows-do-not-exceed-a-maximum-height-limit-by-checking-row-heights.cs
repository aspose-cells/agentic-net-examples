// Title: C# example: enforce a maximum row height after auto‑fitting rows with Aspose.Cells
// AI Prompts: Write a C# method that opens an Excel workbook with Aspose.Cells, calls AutoFitRows on each worksheet, and returns the addresses of rows whose Height property exceeds a specified point limit. | Create a reusable Aspose.Cells utility that iterates all worksheets, auto‑fits rows, checks each row's Height against a configurable maximum, and logs any rows that violate the limit.
// Common Searches: Aspose.Cells C# enforce maximum row height after AutoFitRows | how to list rows taller than 45 points using Aspose.Cells | C# code to validate row height limits in an Excel workbook with Aspose.Cells | detect rows exceeding height threshold after auto‑fit in Aspose.Cells .NET | check row height after AutoFitRows and log oversized rows in C#
// Tags: auto‑fit rows Aspose.Cells | row height limit Aspose.Cells | validate row height C# | detect oversized rows Excel | iterate worksheets Aspose.Cells

using System;
using Aspose.Cells;

// The sample loads an Excel file, auto‑fits all rows on each worksheet, then scans rows up to the last data row. It reports any row whose Height (in points) exceeds a defined maximum (e.g., 50 pts) and finally saves the workbook.
class AutoFitRowValidator
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Define the maximum allowed row height (in points)
        const double maxRowHeight = 50.0; // Example limit

        // Iterate through each worksheet in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Auto-fit all rows in the current worksheet
            sheet.AutoFitRows();

            // Get the maximum row index that contains data
            int maxRow = sheet.Cells.MaxDataRow;

            // Validate each row's height
            for (int rowIndex = 0; rowIndex <= maxRow; rowIndex++)
            {
                Row row = sheet.Cells.Rows[rowIndex];
                double currentHeight = row.Height; // Height is in points

                // Check if the row height exceeds the defined limit
                if (currentHeight > maxRowHeight)
                {
                    Console.WriteLine($"Worksheet '{sheet.Name}', Row {rowIndex + 1} exceeds max height: {currentHeight} pts (limit: {maxRowHeight} pts)");
                }
            }
        }

        // Save the workbook if any modifications were made (optional)
        workbook.Save("output.xlsx");
    }
}
