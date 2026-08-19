// Title: C# – Apply Aspose.Cells AutoFilter on column AJ to show values > 500
// Description: Creates a workbook, adds a header and numeric data to column AJ, defines the AutoFilter range, applies a custom GreaterThan filter with a threshold of 500, refreshes the view, and saves the file as an Excel workbook.
// Keywords: Aspose.Cells | C# AutoFilter | column AJ filter | greater than 500 | custom filter Aspose | set AutoFilter range | filter numeric values | Excel automation
// Common Searches: Aspose.Cells apply AutoFilter to specific column C# | filter rows where column value > 500 using Aspose.Cells | set AutoFilter range and custom criteria in .NET | C# code to show only high‑value rows in Excel with Aspose | how to use GreaterThan filter in Aspose.Cells
// Developer Intent: Show only the rows whose AJ column value exceeds 500 by applying a custom AutoFilter with Aspose.Cells for .NET.
// Use Cases: Financial statements that list transactions above a certain amount. | Data dashboards that need to hide low‑value entries for clarity. | Export routines that require only high‑value records in the output file.
// AI Prompts: Generate C# code using Aspose.Cells to filter column AJ for values greater than 500 and save the workbook. | Explain how to modify the example to filter a different column or change the numeric threshold. | Provide a snippet that removes the AutoFilter and reveals all rows again.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Creates a workbook, adds a header and numeric data to column AJ, defines the AutoFilter range, applies a custom GreaterThan filter with a threshold of 500, refreshes the view, and saves the file as an Excel workbook.
    public class AutoFilterColumnAJGreaterThan500
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // -------------------------------------------------
            // Sample data (header + some rows) for demonstration
            // -------------------------------------------------
            // Header in column AJ (index 35)
            worksheet.Cells[0, 35].PutValue("Amount");

            // Populate rows with sample numeric values
            int[] sampleValues = { 250, 600, 450, 800, 300, 950 };
            for (int i = 0; i < sampleValues.Length; i++)
            {
                // Data starts from row 1 (zero‑based index)
                worksheet.Cells[i + 1, 35].PutValue(sampleValues[i]);
            }

            // -------------------------------------------------
            // Apply AutoFilter to column AJ (index 35)
            // -------------------------------------------------
            // Determine the last row that contains data (zero‑based)
            int lastDataRow = worksheet.Cells.MaxDataRow;

            // Set the filter range: from header row (0) to last data row, column AJ only
            // totalRows = lastDataRow + 1 (including header row)
            worksheet.AutoFilter.SetRange(0, 35, lastDataRow + 1);

            // Apply a custom filter: show only rows where the value > 500
            worksheet.AutoFilter.Custom(35, FilterOperatorType.GreaterThan, 500);

            // Refresh the filter to hide rows that do not meet the criteria
            worksheet.AutoFilter.Refresh();

            // -------------------------------------------------
            // Save the workbook
            // -------------------------------------------------
            string outputPath = "AutoFilter_AJ_GreaterThan500.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved as {outputPath}");
        }
    }
}
