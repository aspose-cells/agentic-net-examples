// Title: Set a print area that includes the header row and freeze the header in an Excel worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells to define a print area covering A1:C20, repeat the first row on every printed page, and freeze the header at cell A2. | Update an existing Aspose.Cells workbook to configure PrintTitleRows for the header row and apply FreezePanes at A2 while keeping the current print area.
// Common Searches: asp.net aspose.cells set print area and freeze header row in C# | c# aspose.cells repeat first row on each printed page and freeze panes | how to use PrintTitleRows with FreezePanes in Aspose.Cells .NET
// Tags: Aspose.Cells set print area C# | Aspose.Cells repeat header rows on print | Aspose.Cells freeze panes at specific cell | Aspose.Cells configure PrintTitleRows .NET | Aspose.Cells worksheet print setup

using System;
using Aspose.Cells;

// Demonstrates creating a workbook, populating data, setting the print area to A1:C20, repeating the first row on every printed page, freezing the header row at A2, and saving the file with Aspose.Cells for .NET.
class PrintAreaAndFreezeDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // ----- Populate sample data -----
        // Header row
        sheet.Cells["A1"].PutValue("ID");
        sheet.Cells["B1"].PutValue("Name");
        sheet.Cells["C1"].PutValue("Amount");

        // Data rows
        for (int i = 2; i <= 20; i++)
        {
            sheet.Cells[$"A{i}"].PutValue(i - 1);
            sheet.Cells[$"B{i}"].PutValue($"Item {i - 1}");
            sheet.Cells[$"C{i}"].PutValue((i - 1) * 10);
        }

        // ----- Set print area that includes the header and all data rows -----
        sheet.PageSetup.PrintArea = "A1:C20";

        // Repeat the header row on each printed page
        sheet.PageSetup.PrintTitleRows = "$1:$1";

        // ----- Freeze the header row -----
        // Freeze at cell A2, freezing 1 row (the header) and 0 columns
        sheet.FreezePanes("A2", 1, 0);

        // Save the workbook
        workbook.Save("PrintAreaAndFreezeDemo.xlsx");
    }
}
