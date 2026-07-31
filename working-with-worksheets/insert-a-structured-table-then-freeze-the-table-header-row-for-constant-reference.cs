// Title: Add a ListObject table and freeze its header row using Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a new workbook, populate cells A1:B4 with product data, convert the range into a styled ListObject table, freeze the first row with FreezePanes, and save the result as TableWithFrozenHeader.xlsx.
// Keywords: Aspose.Cells ListObject C# | freeze header row .NET | Excel table style Aspose | FreezePanes Aspose.Cells example | create structured table Excel C#
// Common Searches: Aspose.Cells add ListObject and freeze first row | C# freeze header in Excel worksheet using Aspose | how to apply table style and freeze panes Aspose.Cells | create Excel table with frozen header .NET
// Developer Intent: Create a styled Excel table and keep its header visible while scrolling.
// Use Cases: Build a sales dashboard where product rows are in a formatted table with a locked header for quick navigation. | Export large inventory lists to Excel, apply a built‑in table style, and freeze column titles for better readability. | Generate dynamic reports in .NET applications that require a ListObject table with a persistent header row.
// AI Prompts: Write C# code with Aspose.Cells that adds a ListObject covering a variable range and freezes the top row. | Show how to set TableStyleType for a ListObject and then apply FreezePanes to keep the header visible. | Explain how to calculate the last data row programmatically and use that value to freeze the header after creating a table.

using Aspose.Cells;
using Aspose.Cells.Tables;

// Demonstrates how to create a new workbook, populate cells A1:B4 with product data, convert the range into a styled ListObject table, freeze the first row with FreezePanes, and save the result as TableWithFrozenHeader.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate data including header row
        sheet.Cells["A1"].PutValue("Product");
        sheet.Cells["B1"].PutValue("Quantity");
        sheet.Cells["A2"].PutValue("Apple");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["A3"].PutValue("Banana");
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["A4"].PutValue("Cherry");
        sheet.Cells["B4"].PutValue(30);

        // Add a structured table (ListObject) covering the range A1:B4 with headers
        int tableIndex = sheet.ListObjects.Add(0, 0, 3, 1, true);
        ListObject table = sheet.ListObjects[tableIndex];
        table.TableStyleType = TableStyleType.TableStyleMedium2;

        // Freeze the header row (first row) so it stays visible while scrolling
        // Freeze panes at row 2 (index 2) to keep row 1 frozen
        sheet.FreezePanes(2, 1, 1, 0);

        // Save the workbook
        workbook.Save("TableWithFrozenHeader.xlsx");
    }
}
