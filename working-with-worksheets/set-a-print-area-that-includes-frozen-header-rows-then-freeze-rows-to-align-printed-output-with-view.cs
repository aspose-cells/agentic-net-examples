// Title: Define Print Area, Repeat Header Row, and Freeze Panes with Aspose.Cells for .NET (C#)
// Description: C# sample creates a workbook, fills A1:C20, sets the print area to A1:C20, repeats row 1 on each printed page, freezes the top row, and saves as PrintAreaAndFreezeDemo.xlsx using Aspose.Cells.
// Keywords: Aspose.Cells | C# | .NET | print area | freeze panes | repeat header row | PageSetup.PrintTitleRows | Worksheet.FreezePanes | Excel export | report printing
// Common Searches: Aspose.Cells set print area C# | freeze first row Aspose.Cells .NET | repeat header rows on printed pages Aspose.Cells | combine PrintTitleRows and FreezePanes Aspose.Cells | C# code for print area and frozen header in Excel
// Developer Intent: Configure a worksheet to print a specific range, repeat the header on each page, and keep the header frozen in the UI.
// Use Cases: Generate paginated reports where column headings appear on every printed page and stay visible while scrolling. | Create invoices or statements that require a fixed header row both in print and on‑screen. | Export large data tables with a defined print area and synchronized on‑screen view for user review. | Automate Excel workbook preparation for dashboards that need consistent header visibility across print and screen.
// AI Prompts: Write C# code using Aspose.Cells to set a print area, define PrintTitleRows, and freeze the top row. | Show how to programmatically determine the used range, assign it as the print area, and apply FreezePanes for the first N rows in Aspose.Cells. | Explain the relationship between PageSetup.PrintTitleRows and Worksheet.FreezePanes and how to use them together to align printed output with the worksheet view. | Provide a step‑by‑step guide for creating a printable Excel report with a frozen header using Aspose.Cells for .NET.

using System;
using Aspose.Cells;

// C# sample creates a workbook, fills A1:C20, sets the print area to A1:C20, repeats row 1 on each printed page, freezes the top row, and saves as PrintAreaAndFreezeDemo.xlsx using Aspose.Cells.
class PrintAreaAndFreezeDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // -------------------------------------------------
        // Populate sample data (header + several rows)
        // -------------------------------------------------
        worksheet.Cells["A1"].PutValue("Header1");
        worksheet.Cells["B1"].PutValue("Header2");
        worksheet.Cells["C1"].PutValue("Header3");

        for (int i = 2; i <= 20; i++)
        {
            worksheet.Cells[$"A{i}"].PutValue($"DataA{i - 1}");
            worksheet.Cells[$"B{i}"].PutValue($"DataB{i - 1}");
            worksheet.Cells[$"C{i}"].PutValue($"DataC{i - 1}");
        }

        // -------------------------------------------------
        // Set the print area to include the header row and data rows
        // -------------------------------------------------
        worksheet.PageSetup.PrintArea = "A1:C20";

        // -------------------------------------------------
        // Repeat the header row on each printed page
        // -------------------------------------------------
        worksheet.PageSetup.PrintTitleRows = "$1:$1";

        // -------------------------------------------------
        // Freeze the header row so the view aligns with the printed output
        // Freeze at cell A2 (row index 2) with 1 frozen row and 0 frozen columns
        // -------------------------------------------------
        worksheet.FreezePanes("A2", 1, 0);

        // -------------------------------------------------
        // Save the workbook
        // -------------------------------------------------
        workbook.Save("PrintAreaAndFreezeDemo.xlsx");
    }
}
