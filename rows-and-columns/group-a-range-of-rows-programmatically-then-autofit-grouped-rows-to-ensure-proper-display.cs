// Title: Group rows and auto‑fit height using Aspose.Cells for .NET (C#)
// Description: This C# sample builds a new workbook, writes ten rows of test data, groups rows 2‑6 while keeping them visible, sets the summary line above the detail rows, automatically resizes the grouped rows to show their content, and writes the file as GroupRowsAutoFitDemo.xlsx.
// Keywords: Aspose.Cells | C# | .NET | group rows | auto fit rows | outline summary row | Worksheet.AutoFitRows | Cells.GroupRows | Excel row grouping programmatically | Excel automation with Aspose
// Common Searches: how to group rows in Aspose.Cells C# | auto fit grouped rows Aspose.Cells .NET | set summary row above detail rows Aspose.Cells | C# example for worksheet outline and auto‑fit | Aspose.Cells GroupRows and AutoFitRows usage
// Developer Intent: Create a collapsible block of rows and automatically adjust its height so all cell values are fully visible.
// Use Cases: Design a drill‑down report where each section can be expanded and the row height adapts to wrapped text. | Generate an invoice that groups line items by category and ensures description fields are not truncated. | Build a hierarchical data view in Excel where each grouped segment is displayed with optimal row spacing.
// AI Prompts: Provide C# code that groups rows 5‑12 in an Aspose.Cells worksheet, places the summary row below the details, and calls AutoFitRows for that range. | Show an Aspose.Cells snippet that hides outline symbols, groups rows, and adjusts both column widths and row heights for better readability. | Explain the effect of Outline.SummaryRowBelow on grouped rows and how to combine it with Worksheet.AutoFitRows for proper display.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // This C# sample builds a new workbook, writes ten rows of test data, groups rows 2‑6 while keeping them visible, sets the summary line above the detail rows, automatically resizes the grouped rows to show their content, and writes the file as GroupRowsAutoFitDemo.xlsx.
    public class GroupRowsAndAutoFitDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Populate sample data in rows 0 to 9
                for (int i = 0; i < 10; i++)
                {
                    cells[i, 0].PutValue($"Item {i + 1}");
                    cells[i, 1].PutValue($"Description for item {i + 1}");
                    cells[i, 2].PutValue(i * 10);
                }

                // Group rows 2 through 6 (zero‑based index) and keep them visible
                cells.GroupRows(2, 6, false);

                // Set the summary row position (false = above the detail rows)
                worksheet.Outline.SummaryRowBelow = false;

                // Auto‑fit the height of the grouped rows to display their content properly
                worksheet.AutoFitRows(2, 6);

                // Save the workbook
                workbook.Save("GroupRowsAutoFitDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            GroupRowsAndAutoFitDemo.Run();
        }
    }
}
