// Title: C# – Set Row Height and Freeze Top Rows with Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, assign custom heights (30 pt and 40 pt) to the first two rows, add sample text, freeze those rows using Worksheet.FreezePanes, and save the file as RowHeightsAndFreeze.xlsx.
// Keywords: Aspose.Cells C# | set row height Aspose.Cells | custom row height Excel .NET | freeze panes top rows | Worksheet.FreezePanes example | Excel row height points | freeze first rows Aspose.Cells | C# Excel automation | Aspose.Cells row formatting | freeze top rows C#
// Common Searches: Aspose.Cells set row height C# | how to freeze first rows with Aspose.Cells | C# freeze panes after setting row height | custom row height Excel using Aspose.Cells .NET | Worksheet.FreezePanes syntax C# | freeze top two rows Aspose.Cells example
// Developer Intent: Apply custom heights to the initial rows and lock them in view while scrolling the worksheet.
// Use Cases: Design a report header with larger rows that stay visible during navigation. | Generate spreadsheets where title rows need distinct heights and remain fixed. | Create printable sheets with top rows of specific height that are frozen for easy reference.
// AI Prompts: Show C# code to set row heights in points and freeze the first N rows with Aspose.Cells. | Provide an Aspose.Cells example that freezes rows only, preserving custom row height formatting. | Explain the parameters of Worksheet.FreezePanes when freezing rows in a .NET workbook.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, assign custom heights (30 pt and 40 pt) to the first two rows, add sample text, freeze those rows using Worksheet.FreezePanes, and save the file as RowHeightsAndFreeze.xlsx.
    public class SetRowHeightsAndFreezeDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Set custom heights for the first two rows (index 0 and 1)
                // Height is in points (1 point = 1/72 inch)
                cells.SetRowHeight(0, 30); // Row 1 height = 30 points
                cells.SetRowHeight(1, 40); // Row 2 height = 40 points

                // Add sample data to visualize the rows
                cells["A1"].PutValue("First row with custom height");
                cells["A2"].PutValue("Second row with custom height");

                // Freeze the first two rows so that their custom heights stay visible while scrolling
                // Freeze at the cell just below the rows to be frozen (row index 2 -> third row)
                // No columns are frozen, so freezedColumns = 0
                worksheet.FreezePanes(2, 0, 2, 0);
                // Equivalent call using cell name:
                // worksheet.FreezePanes("A3", 2, 0);

                // Save the workbook
                workbook.Save("RowHeightsAndFreeze.xlsx");
                Console.WriteLine("Workbook saved as RowHeightsAndFreeze.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            SetRowHeightsAndFreezeDemo.Run();
        }
    }
}
