// Title: Hide Table Header Row in Aspose.Cells for .NET – Compact Dashboard Layout (C#)
// Description: Creates a workbook, adds sample data, defines a ListObject (Excel table), applies a built‑in style, disables the header row with ShowHeaderRow = false, and saves the file as DashboardCompactTable.xlsx for a space‑saving dashboard view.
// Keywords: Aspose.Cells C# | Hide table header | ShowHeaderRow false | compact dashboard Excel | ListObject header visibility | .NET Excel table example | Aspose.Cells table styling
// Common Searches: Aspose.Cells hide table header row C# | ShowHeaderRow property false example | compact Excel dashboard table Aspose | remove ListObject header Aspose.Cells | C# code to hide Excel table header
// Developer Intent: Turn off the table header to produce a tighter, dashboard‑friendly layout.
// Use Cases: Design a KPI dashboard where vertical space is limited and the header row is unnecessary. | Generate printable reports that omit redundant column titles for a cleaner look. | Create Excel files for downstream PDF conversion where the header would be duplicated elsewhere.
// AI Prompts: Write C# code using Aspose.Cells to add a ListObject and hide its header row for a compact layout. | Explain the impact of the ShowHeaderRow property on table appearance and file size in Aspose.Cells. | Show how to toggle the header row of an existing Aspose.Cells table at runtime.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsDashboard
{
    // Creates a workbook, adds sample data, defines a ListObject (Excel table), applies a built‑in style, disables the header row with ShowHeaderRow = false, and saves the file as DashboardCompactTable.xlsx for a space‑saving dashboard view.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data for the table
            worksheet.Cells["A1"].PutValue("Product");
            worksheet.Cells["B1"].PutValue("Quantity");
            worksheet.Cells["A2"].PutValue("Apple");
            worksheet.Cells["B2"].PutValue(50);
            worksheet.Cells["A3"].PutValue("Banana");
            worksheet.Cells["B3"].PutValue(30);
            worksheet.Cells["A4"].PutValue("Cherry");
            worksheet.Cells["B4"].PutValue(20);

            // Add a ListObject (table) covering the data range (including header row)
            // Parameters: firstRow, firstColumn, totalRows, totalColumns, hasHeaders
            int tableIndex = worksheet.ListObjects.Add(0, 0, 3, 1, true);
            ListObject table = worksheet.ListObjects[tableIndex];

            // Apply a built‑in table style (optional, for visual consistency)
            table.TableStyleType = TableStyleType.TableStyleMedium2;

            // Hide the header row for a compact dashboard view
            table.ShowHeaderRow = false;

            // Save the workbook to a file
            workbook.Save("DashboardCompactTable.xlsx");
        }
    }
}
