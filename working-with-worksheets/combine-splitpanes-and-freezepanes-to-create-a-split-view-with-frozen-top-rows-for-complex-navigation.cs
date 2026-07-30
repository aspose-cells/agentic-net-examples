// Title: Aspose.Cells .NET: Split Pane and Freeze Top Rows for a Fixed‑Header View (C#)
// Description: This example creates a new workbook, populates 100 rows of sample data, splits the worksheet window into panes, freezes the first five rows, sets the bottom pane to start after the frozen area, and saves the result as SplitAndFreezeDemo.xlsx.
// Keywords: Aspose.Cells C# split panes | Aspose.Cells FreezePanes | split and freeze rows .NET | fixed header worksheet Aspose | first visible row bottom pane | Excel API split view | worksheet navigation Aspose.Cells | C# Excel pane scrolling
// Common Searches: how to split worksheet and freeze header rows using Aspose.Cells | Aspose.Cells split view with frozen top rows example | set first visible row of bottom pane after FreezePanes Aspose | C# code to combine Split() and FreezePanes in Aspose.Cells | freeze top rows while keeping split panes in Excel file
// Developer Intent: Create a worksheet that shows a split view with the top five rows frozen for constant header visibility.
// Use Cases: Display large tables where column headers stay visible while scrolling the data section. | Design a reporting sheet with a fixed header and independent scrolling panes for detailed rows. | Build a dashboard where the upper rows act as navigation controls and the lower pane shows drill‑down information.
// AI Prompts: Write C# code with Aspose.Cells to split a worksheet at row 8, freeze the first 3 rows, and set the bottom pane's first visible row to 4. | Explain the interaction between Split() and FreezePanes() parameters when creating a split view with a frozen header in Aspose.Cells. | Provide a step‑by‑step tutorial for adjusting the first visible row of the bottom pane after freezing rows in an Aspose.Cells workbook.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // This example creates a new workbook, populates 100 rows of sample data, splits the worksheet window into panes, freezes the first five rows, sets the bottom pane to start after the frozen area, and saves the result as SplitAndFreezeDemo.xlsx.
    public class SplitAndFreezeDemo
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
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate the worksheet with sample data to demonstrate navigation
            for (int row = 0; row < 100; row++)
            {
                for (int col = 0; col < 10; col++)
                {
                    sheet.Cells[row, col].Value = $"R{row + 1}C{col + 1}";
                }
            }

            // Split the window into panes (default split position is at the middle of the view)
            sheet.Split();

            // Freeze the top 5 rows while keeping the split panes
            // Parameters: row index, column index, number of frozen rows, number of frozen columns
            sheet.FreezePanes(5, 0, 5, 0);

            // Optionally adjust the first visible row of the bottom pane for smoother scrolling
            PaneCollection panes = sheet.GetPanes();
            panes.FirstVisibleRowOfBottomPane = 5; // start bottom pane after the frozen rows

            // Save the workbook
            workbook.Save("SplitAndFreezeDemo.xlsx");
            Console.WriteLine("Workbook saved as SplitAndFreezeDemo.xlsx");
        }
    }
}
