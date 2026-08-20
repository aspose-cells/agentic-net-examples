// Title: Add an Excel table (A1:C10) with filter buttons using Aspose.Cells for .NET
// Description: Creates a new workbook, fills range A1:C10 with headers and sample data, adds a ListObject that spans the range, enables AutoFilter with visible filter buttons on the header row, and saves the file as TableWithFilters.xlsx.
// Keywords: Aspose.Cells C# table from range | ListObject AutoFilter Aspose | Excel filter buttons .NET | create Excel table programmatically | Aspose.Cells enable filter dropdown
// Common Searches: Aspose.Cells add ListObject with filter buttons | C# enable AutoFilter on Excel table using Aspose | how to create table A1:C10 in Aspose.Cells | show filter dropdowns in Aspose.Cells worksheet | Aspose.Cells example for Excel table with filters
// Developer Intent: Create a table covering A1:C10 and turn on filter buttons for the header row in an Excel file.
// Use Cases: Export a report that lets end users sort and filter data directly in Excel. | Provide a pre‑formatted data template with interactive filter controls. | Generate spreadsheets with consistent table styling and built‑in AutoFilter for analysis.
// AI Prompts: Generate C# code with Aspose.Cells that adds a ListObject for A1:C10, enables AutoFilter, and shows filter buttons on the header row. | Show how to populate sample data, create an Excel table, activate filter dropdowns, and save the workbook as .xlsx using Aspose.Cells. | Explain how to apply a custom table style while keeping filter buttons visible in Aspose.Cells for .NET.

using Aspose.Cells;
using Aspose.Cells.Tables;

// Creates a new workbook, fills range A1:C10 with headers and sample data, adds a ListObject that spans the range, enables AutoFilter with visible filter buttons on the header row, and saves the file as TableWithFilters.xlsx.
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
            worksheet.Cells[row - 1, 1].PutValue(row * 10);
            worksheet.Cells[row - 1, 2].PutValue(row * 100);
        }

        // Add a ListObject (table) that covers the range A1:C10 (rows 0‑9, columns 0‑2)
        int tableIndex = worksheet.ListObjects.Add(0, 0, 9, 2, true);
        ListObject table = worksheet.ListObjects[tableIndex];

        // Enable AutoFilter for the table and show filter buttons on the header row
        table.HasAutoFilter = true;
        table.AutoFilter.ShowFilterButton = true;

        // Save the workbook
        workbook.Save("TableWithFilters.xlsx");
    }
}
