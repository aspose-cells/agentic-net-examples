// Title: C# Split Worksheet into Panes and Freeze Top Rows with Aspose.Cells
// Description: Demonstrates how to create a new workbook, populate 100 rows, split the worksheet into four panes, freeze the first five rows, set the bottom pane to start at row 11, and save the result as SplitAndFreezeDemo.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# split panes | Aspose.Cells freeze panes | split view frozen header .NET | FreezePanes with Split in Aspose.Cells | worksheet split and freeze example | C# Excel pane manipulation
// Common Searches: Aspose.Cells split worksheet into panes C# | freeze top rows while keeping split panes Aspose.Cells | set first visible row of bottom pane Aspose.Cells | how to combine Split and FreezePanes in .NET | split view with frozen header Aspose.Cells
// Developer Intent: Create a worksheet that shows a split view and keeps the top rows fixed for easier navigation in a large dataset.
// Use Cases: Display large tables with a frozen header while allowing independent scrolling of lower sections. | Build Excel dashboards where summary rows remain visible across multiple panes. | Navigate to a specific data block by positioning the bottom pane at a chosen row after freezing the header.
// AI Prompts: Write C# code using Aspose.Cells to split a worksheet into four panes, freeze the first five rows, and set the bottom pane to start at row 11. | Explain the FreezePanes parameters and how they interact with Split in Aspose.Cells. | Provide a step‑by‑step tutorial for adjusting the first visible row of the bottom pane after splitting and freezing.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a new workbook, populate 100 rows, split the worksheet into four panes, freeze the first five rows, set the bottom pane to start at row 11, and save the result as SplitAndFreezeDemo.xlsx using Aspose.Cells for .NET.
    public class SplitAndFreezeDemo
    {
        // Entry point required for console application
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
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate the worksheet with sample data to demonstrate scrolling
            for (int row = 0; row < 100; row++)
            {
                for (int col = 0; col < 10; col++)
                {
                    sheet.Cells[row, col].Value = $"R{row + 1}C{col + 1}";
                }
            }

            // Split the window into four panes (no parameters needed)
            sheet.Split();

            // Freeze the top 5 rows while keeping the split panes
            // Parameters: row index, column index, number of frozen rows, number of frozen columns
            // Here we freeze rows 0‑4 (5 rows) and no columns
            sheet.FreezePanes(5, 0, 5, 0);

            // Optionally adjust the visible part of the bottom pane
            PaneCollection panes = sheet.GetPanes();
            panes.FirstVisibleRowOfBottomPane = 10; // start showing row 11 in the bottom pane

            // Save the workbook
            string outputPath = "SplitAndFreezeDemo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
    }
}
