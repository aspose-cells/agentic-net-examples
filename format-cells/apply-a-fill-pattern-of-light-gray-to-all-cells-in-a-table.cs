// Title: Apply Light Gray Fill Pattern to All Cells of a ListObject Table with Aspose.Cells for .NET (C#)
// Description: This C# example creates a workbook, adds a ListObject covering A1:C4, defines a style using the Gray25 pattern and LightGray foreground, applies the style to every cell in the table, and saves the file as TableWithLightGrayFill.xlsx.
// Keywords: Aspose.Cells C# fill pattern | light gray background Aspose.Cells | Gray25 pattern | ListObject style | apply style to table cells | Aspose.Cells table formatting | C# Excel cell background | Aspose.Cells SetStyle | Excel table background color | Aspose.Cells range styling
// Common Searches: Aspose.Cells set light gray fill for table cells C# | How to apply Gray25 pattern to ListObject in Aspose.Cells | C# Aspose.Cells change background color of entire table | Apply style to all cells of a ListObject Aspose.Cells | Excel table fill pattern using Aspose.Cells .NET
// Developer Intent: Developer wants to color every cell of a ListObject table with a light gray fill (Gray25) using Aspose.Cells in C#.
// Use Cases: Generate printable reports where the data table has a uniform light gray background. | Visually separate a table from surrounding content in an automated Excel export. | Prepare a workbook for further conditional formatting by first applying a base gray fill to all table cells. | Create a consistent theme for tables across multiple sheets generated programmatically.
// AI Prompts: Write C# code using Aspose.Cells to apply a Gray25 light gray fill pattern to all cells of an existing ListObject. | Show how to set a light gray background for a ListObject without looping through each cell, using the Range object. | Explain how to change the fill pattern to LightHorizontal and adjust the foreground color while keeping the same table styling approach. | Provide a step‑by‑step guide to create a style, assign it to a table range, and save the workbook in Aspose.Cells .NET.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Tables;

// This C# example creates a workbook, adds a ListObject covering A1:C4, defines a style using the Gray25 pattern and LightGray foreground, applies the style to every cell in the table, and saves the file as TableWithLightGrayFill.xlsx.
class ApplyLightGrayFillToTable
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the table (A1:C4)
        sheet.Cells["A1"].PutValue("Product");
        sheet.Cells["B1"].PutValue("Category");
        sheet.Cells["C1"].PutValue("Price");

        sheet.Cells["A2"].PutValue("Apple");
        sheet.Cells["B2"].PutValue("Fruit");
        sheet.Cells["C2"].PutValue(1.2);

        sheet.Cells["A3"].PutValue("Carrot");
        sheet.Cells["B3"].PutValue("Vegetable");
        sheet.Cells["C3"].PutValue(0.8);

        sheet.Cells["A4"].PutValue("Bread");
        sheet.Cells["B4"].PutValue("Bakery");
        sheet.Cells["C4"].PutValue(2.5);

        // Add a ListObject (table) covering the data range
        int tableIndex = sheet.ListObjects.Add(0, 0, 4, 2, true);
        ListObject table = sheet.ListObjects[tableIndex];
        table.ShowHeaderRow = true;
        table.ShowTableStyleFirstColumn = true;
        table.ShowTableStyleLastColumn = true;
        table.ShowTableStyleRowStripes = true;
        table.ShowTableStyleColumnStripes = true;

        // Create a style with a light gray fill pattern
        Style lightGrayStyle = workbook.CreateStyle();
        lightGrayStyle.Pattern = BackgroundType.Gray25;          // Light gray pattern
        lightGrayStyle.ForegroundColor = Color.LightGray;       // Light gray foreground
        lightGrayStyle.BackgroundColor = Color.White;           // Optional background color

        // Apply the style to every cell inside the table
        for (int row = table.StartRow; row <= table.EndRow; row++)
        {
            for (int col = table.StartColumn; col <= table.EndColumn; col++)
            {
                Cell cell = sheet.Cells[row, col];
                cell.SetStyle(lightGrayStyle);
            }
        }

        // Save the workbook
        workbook.Save("TableWithLightGrayFill.xlsx", SaveFormat.Xlsx);
    }
}
