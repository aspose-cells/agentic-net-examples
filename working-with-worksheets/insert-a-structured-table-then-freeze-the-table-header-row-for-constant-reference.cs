// Title: C# – Add a Structured Table and Freeze Its Header Row with Aspose.Cells for .NET
// Description: This example creates a new workbook, writes header and sample data to A1:C3, inserts a ListObject table with a built‑in style, freezes the first row (header) using FreezePanes at cell A2, and saves the result as TableWithFrozenHeader.xlsx.
// Keywords: Aspose.Cells C# table example | Add ListObject Aspose.Cells | Freeze header row Aspose.Cells | FreezePanes C# Aspose.Cells | styled Excel table .NET | Aspose.Cells sample code | GitHub Aspose.Cells worksheet
// Common Searches: how to add a ListObject table in Aspose.Cells for .NET | freeze first row in Excel using Aspose.Cells C# | Aspose.Cells example for table style and FreezePanes | C# code to create a structured table and freeze header | Aspose.Cells tutorial freeze panes with table
// Developer Intent: Generate an Excel worksheet, convert a range into a styled ListObject table, and keep the header row visible while scrolling.
// Use Cases: Building financial reports where column titles must stay in view for large data sets. | Exporting dashboard metrics to Excel with automatic table formatting and a frozen top row. | Creating data‑entry templates that apply a table style and lock the header for easier navigation.
// AI Prompts: Show C# code to insert a ListObject table into a range and freeze the header row with Aspose.Cells. | Provide an Aspose.Cells example that applies a medium table style and uses FreezePanes at cell A2. | Explain the interaction between ListObject tables and FreezePanes in Aspose.Cells for .NET, including any constraints.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

// This example creates a new workbook, writes header and sample data to A1:C3, inserts a ListObject table with a built‑in style, freezes the first row (header) using FreezePanes at cell A2, and saves the result as TableWithFrozenHeader.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data (including header row)
        worksheet.Cells["A1"].PutValue("Header1");
        worksheet.Cells["B1"].PutValue("Header2");
        worksheet.Cells["C1"].PutValue("Header3");

        worksheet.Cells["A2"].PutValue(10);
        worksheet.Cells["B2"].PutValue(20);
        worksheet.Cells["C2"].PutValue(30);

        worksheet.Cells["A3"].PutValue(40);
        worksheet.Cells["B3"].PutValue(50);
        worksheet.Cells["C3"].PutValue(60);

        // Add a structured table (ListObject) covering the range A1:C3 with headers
        int tableIndex = worksheet.ListObjects.Add(0, 0, 2, 2, true);
        ListObject table = worksheet.ListObjects[tableIndex];
        table.TableStyleType = TableStyleType.TableStyleMedium2; // optional styling

        // Freeze the header row so it stays visible while scrolling
        // Freeze at cell A2 (row index 1) with 1 frozen row and 0 frozen columns
        worksheet.FreezePanes("A2", 1, 0);

        // Save the workbook
        workbook.Save("TableWithFrozenHeader.xlsx");
    }
}
