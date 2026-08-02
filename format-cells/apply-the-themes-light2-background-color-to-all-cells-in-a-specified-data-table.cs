// Title: Apply Theme Light2 Background to an Excel Table with Aspose.Cells for .NET (C#)
// Description: This C# example creates a new workbook, fills range A1:C5, converts it into a ListObject (Excel table), assigns the built‑in TableStyleLight2 (which uses the theme’s Light2 background color) to the entire table, and saves the file as TableWithLight2Style.xlsx.
// Keywords: Aspose.Cells | C# Excel table style | TableStyleLight2 | theme Light2 background | ListObject styling | apply built‑in table style programmatically | Excel theme colors .NET | Aspose.Cells TableStyleType
// Common Searches: Aspose.Cells set table style Light2 | C# apply theme background to Excel table | How to use TableStyleLight2 with ListObject | Change Excel table style programmatically .NET | Apply built‑in table style Aspose.Cells
// Developer Intent: Set the built‑in Light2 style so every cell in a specified ListObject uses the theme’s Light2 background color.
// Use Cases: Generate a report where all tables share a consistent Light2 theme for branding. | Update existing worksheets to a uniform theme‑based background without manual formatting. | Create multiple tables in a single workbook, each automatically styled with Light2 for visual cohesion.
// AI Prompts: Write C# code using Aspose.Cells to change an existing ListObject’s TableStyleType to TableStyleLight2 and save the workbook. | Show how to create a new workbook, add data to A1:D10, convert the range to a table, and apply the Light2 built‑in style with Aspose.Cells for .NET. | Provide a snippet that updates a previously generated Excel table to use the theme’s Light2 background color via TableStyleType in C#.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

// This C# example creates a new workbook, fills range A1:C5, converts it into a ListObject (Excel table), assigns the built‑in TableStyleLight2 (which uses the theme’s Light2 background color) to the entire table, and saves the file as TableWithLight2Style.xlsx.
class ApplyLight2Style
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Populate sample data for the table (range A1:C5)
        cells["A1"].PutValue("Header1");
        cells["B1"].PutValue("Header2");
        cells["C1"].PutValue("Header3");
        for (int row = 2; row <= 5; row++)
        {
            cells[$"A{row}"].PutValue($"R{row - 1}C1");
            cells[$"B{row}"].PutValue($"R{row - 1}C2");
            cells[$"C{row}"].PutValue($"R{row - 1}C3");
        }

        // Create a ListObject (Excel table) that covers the data range
        int tableIndex = worksheet.ListObjects.Add("A1", "C5", true);
        ListObject table = worksheet.ListObjects[tableIndex];

        // Apply the built‑in Light2 table style, which uses the theme's Light2 background color
        table.TableStyleType = TableStyleType.TableStyleLight2;

        // Save the workbook
        workbook.Save("TableWithLight2Style.xlsx");
    }
}
