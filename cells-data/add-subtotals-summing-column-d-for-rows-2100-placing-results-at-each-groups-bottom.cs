// Title: C# Aspose.Cells – Insert Subtotal Rows Summing Column D Grouped by Column C (Rows 2‑100)
// Description: This example creates a workbook, fills rows 2‑100 with group identifiers in column C and numeric values in column D, then uses the Cells.Subtotal method to add a subtotal row below each group that sums column D. The file is saved as SubtotalResult.xlsx.
// Keywords: Aspose.Cells subtotal C# | add subtotal rows Excel .NET | group by column C sum column D | Cells.Subtotal example | C# generate Excel subtotals | Aspose.Cells grouping and sum | Excel automation subtotal function | US developers Aspose.Cells tutorial | global Excel subtotal code
// Common Searches: Aspose.Cells how to add subtotals by group | C# subtotal rows for each category in Excel | Cells.Subtotal method example | sum column D for each group in column C using Aspose | create subtotal rows below groups with Aspose.Cells
// Developer Intent: Insert subtotal rows that calculate the sum of column D for each distinct value in column C across rows 2‑100.
// Use Cases: Financial statements that need category‑wise expense totals automatically inserted. | Sales dashboards where regional totals are displayed beneath each region’s data. | Invoice consolidation reports that summarize amounts per client without manual formulas.
// AI Prompts: Show how to modify the Subtotal call to also compute the average of column D per group. | Provide code to style the generated subtotal rows with bold text and a light‑gray fill. | Explain how to add page breaks between groups while keeping the subtotal row directly below each group.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // This example creates a workbook, fills rows 2‑100 with group identifiers in column C and numeric values in column D, then uses the Cells.Subtotal method to add a subtotal row below each group that sums column D. The file is saved as SubtotalResult.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Populate sample data in columns A‑D for rows 2‑100
                for (int row = 1; row < 100; row++) // zero‑based index: row 1 = Excel row 2
                {
                    // Example grouping column (C)
                    cells[row, 2].PutValue(row % 2 == 0 ? "Group1" : "Group2"); // column C
                    // Values to be summed in column D
                    cells[row, 3].PutValue(row * 10);
                }

                // Define the range that includes rows 2‑100 (zero‑based 1‑99) and columns A‑D (0‑3)
                CellArea area = CellArea.CreateCellArea(1, 0, 99, 3);

                // Add subtotals:
                // - Group by column C (zero‑based index 2)
                // - Use SUM function
                // - Apply subtotal to column D (zero‑based index 3)
                // - Replace existing subtotals, no page breaks, place summary below each group
                cells.Subtotal(
                    area,
                    2,                                 // groupBy column C
                    ConsolidationFunction.Sum,         // sum function
                    new int[] { 3 },                   // subtotal column D
                    true,                              // replace existing subtotals
                    false,                             // no page breaks between groups
                    true                               // place summary row below each group
                );

                // Determine output path and ensure directory exists
                string outputPath = "SubtotalResult.xlsx";
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
