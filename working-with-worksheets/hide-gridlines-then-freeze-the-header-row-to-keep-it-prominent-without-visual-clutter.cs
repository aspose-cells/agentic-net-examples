// Title: Aspose.Cells .NET – Hide Gridlines & Freeze Header Row in an Excel Worksheet
// Description: Creates a new workbook, hides worksheet gridlines, inserts a header in A1 with sample data, freezes the first row using Worksheet.FreezePanes, and saves the file as HideGridlinesAndFreezeHeader.xlsx.
// Keywords: Aspose.Cells | C# | .NET | hide gridlines | freeze panes | freeze header row | Excel formatting | worksheet gridlines | freeze top row | Excel report generation
// Common Searches: Aspose.Cells hide gridlines C# | freeze first row Aspose.Cells .NET | how to remove gridlines and lock header in Excel using Aspose | C# example for Worksheet.FreezePanes | Aspose.Cells hide worksheet gridlines and freeze top row
// Developer Intent: Remove worksheet gridlines and keep the header row visible while scrolling.
// Use Cases: Produce a clean, printable report without gridlines and with a fixed header. | Export data to Excel where end‑users need a static top row for navigation. | Generate dashboards that require a clutter‑free view and persistent column titles.
// AI Prompts: Show C# code to hide gridlines and freeze the top row in an existing workbook with Aspose.Cells. | Give an Aspose.Cells example that adds a header, disables gridlines, and freezes the header row. | Explain the parameters of Worksheet.FreezePanes for freezing the first row in Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Creates a new workbook, hides worksheet gridlines, inserts a header in A1 with sample data, freezes the first row using Worksheet.FreezePanes, and saves the file as HideGridlinesAndFreezeHeader.xlsx.
    public class HideGridlinesAndFreezeHeaderDemo
    {
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
            Worksheet worksheet = workbook.Worksheets[0];

            // Hide gridlines
            worksheet.IsGridlinesVisible = false;

            // Add a header row and some sample data
            worksheet.Cells["A1"].PutValue("Header");
            for (int i = 2; i <= 10; i++)
            {
                worksheet.Cells["A" + i].PutValue("Data " + (i - 1));
            }

            // Freeze the header row (first row) at cell A2
            worksheet.FreezePanes(1, 0, 1, 0);

            // Save the workbook
            workbook.Save("HideGridlinesAndFreezeHeader.xlsx");
        }
    }
}
