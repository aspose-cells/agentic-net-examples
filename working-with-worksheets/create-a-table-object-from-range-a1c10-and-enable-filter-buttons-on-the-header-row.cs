// Title: C# – Create an Excel table (A1:C10) with filter buttons using Aspose.Cells
// Description: Demonstrates how to populate cells A1:C10, add a ListObject that spans this range, enable AutoFilter buttons on the header row, and save the workbook as TableWithFilterButtons.xlsx with Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# Excel table | ListObject | AutoFilter | filter buttons | A1:C10 | create table Aspose.Cells | enable filter Aspose.Cells | Excel automation .NET | Aspose.Cells .NET example
// Common Searches: Aspose.Cells add table with filter buttons | C# create Excel table from range A1:C10 | How to enable AutoFilter in Aspose.Cells | ListObject filter buttons C# Aspose | Create Excel table and show filter dropdowns Aspose.Cells
// Developer Intent: Add a ListObject covering A1:C10 and turn on filter buttons for its header row.
// Use Cases: Generate a report workbook where users can instantly sort and filter columns. | Prepare a data‑entry sheet with a predefined table structure and built‑in filter controls. | Export database query results to Excel with an auto‑created table that includes column filters.
// AI Prompts: Show C# code that creates a ListObject for range A1:C10 and enables filter buttons using Aspose.Cells. | Explain how to set the ShowFilterButton property on a table header in Aspose.Cells for .NET. | Provide a step‑by‑step example of populating a range, adding a table, and activating AutoFilter buttons with Aspose.Cells.

using Aspose.Cells;
using Aspose.Cells.Tables;

// Demonstrates how to populate cells A1:C10, add a ListObject that spans this range, enable AutoFilter buttons on the header row, and save the workbook as TableWithFilterButtons.xlsx with Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Optional: populate the range A1:C10 with sample data
        worksheet.Cells["A1"].PutValue("Header1");
        worksheet.Cells["B1"].PutValue("Header2");
        worksheet.Cells["C1"].PutValue("Header3");
        for (int row = 2; row <= 10; row++)
        {
            worksheet.Cells[row - 1, 0].PutValue($"R{row - 1}C1");
            worksheet.Cells[row - 1, 1].PutValue($"R{row - 1}C2");
            worksheet.Cells[row - 1, 2].PutValue($"R{row - 1}C3");
        }

        // Add a ListObject (table) that covers the range A1:C10
        // startRow = 0 (A1), startColumn = 0 (A), endRow = 9 (row 10), endColumn = 2 (C), hasHeaders = true
        int tableIndex = worksheet.ListObjects.Add(0, 0, 9, 2, true);
        ListObject table = worksheet.ListObjects[tableIndex];

        // Enable filter buttons on the header row of the table
        table.AutoFilter.ShowFilterButton = true;

        // Save the workbook
        workbook.Save("TableWithFilterButtons.xlsx");
    }
}
