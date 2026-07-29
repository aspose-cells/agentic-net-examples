// Title: Group rows and auto‑fit their heights in Excel with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a workbook, populate sample data, programmatically group rows 2‑6 (indices 1‑5) while keeping them visible, auto‑fit the grouped rows to the cell content, and save the file as GroupRowsAutoFitDemo.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# group rows | Worksheet.AutoFitRows example | programmatically group Excel rows | auto‑fit row height Aspose | Aspose.Cells .NET row grouping | Excel row height adjustment C#
// Common Searches: Aspose.Cells how to group rows and auto‑fit | C# group rows 2 to 6 Excel Aspose | auto‑fit grouped rows Aspose.Cells .NET | group rows keep visible Aspose.Cells | Worksheet.AutoFitRows after grouping rows
// Developer Intent: The developer needs to group a specific range of rows in an Excel worksheet and automatically adjust their heights so that all cell content is fully visible.
// Use Cases: Generating reports where related rows are collapsed/expanded but still require proper height for long text. | Creating invoices with line‑item sections grouped together while preserving readability of description fields. | Exporting data blocks from a database into Excel, grouping them for hierarchy and applying AutoFitRows to maintain consistent formatting.
// AI Prompts: Generate C# code that groups rows 10‑15 in an existing workbook, auto‑fits their heights, and saves the file as 'Grouped.xlsx' using Aspose.Cells. | Explain the effect of Worksheet.AutoFitRows on grouped rows and list the required parameters for the method in Aspose.Cells for .NET. | Provide a step‑by‑step tutorial for grouping a row range, keeping it visible, and applying AutoFitRows to an existing Excel file with Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, populate sample data, programmatically group rows 2‑6 (indices 1‑5) while keeping them visible, auto‑fit the grouped rows to the cell content, and save the file as GroupRowsAutoFitDemo.xlsx using Aspose.Cells for .NET.
    public class GroupRowsAndAutoFitDemo
    {
        // Entry point for the application
        public static void Main(string[] args)
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
            Cells cells = worksheet.Cells;

            // Populate sample data (optional, just to see the effect)
            for (int i = 0; i < 10; i++)
            {
                cells[i, 0].PutValue($"Row {i + 1}");
                cells[i, 1].PutValue($"Some long text that may require row height adjustment for row {i + 1}");
            }

            // Group rows 2 to 6 (zero‑based indices 1 to 5) and keep them visible
            cells.GroupRows(1, 5, false);

            // Auto‑fit the grouped rows to adjust their heights based on content
            worksheet.AutoFitRows(1, 5);

            // Save the workbook
            workbook.Save("GroupRowsAutoFitDemo.xlsx");
            Console.WriteLine("Workbook saved as GroupRowsAutoFitDemo.xlsx");
        }
    }
}
