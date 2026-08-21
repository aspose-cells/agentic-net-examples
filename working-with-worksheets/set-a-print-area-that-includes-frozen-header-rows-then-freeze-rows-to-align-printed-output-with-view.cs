// Title: Define Print Area, Repeat Header Row, and Freeze Panes with Aspose.Cells for .NET
// Description: Learn how to set a print area (A1:B20), repeat the first row on every printed page, and freeze that header row in an Aspose.Cells workbook using C#.
// Keywords: Aspose.Cells | .NET | C# | print area | repeat header rows | PrintTitleRows | freeze panes | FreezePanes | PageSetup | Excel export | worksheet view alignment
// Common Searches: Aspose.Cells set print area C# | repeat header rows Aspose.Cells .NET | freeze first row Aspose.Cells | align printed view with frozen panes Aspose.Cells | PageSetup.PrintTitleRows example
// Developer Intent: Configure a print area, repeat the header row on each page, and freeze that header to keep the worksheet view and printed output synchronized.
// Use Cases: Create multi‑page reports where column headers appear on every printed page and stay visible while scrolling. | Generate invoices or data sheets with a fixed header that matches the defined print range. | Export large data tables to Excel with a specific print area and synchronized on‑screen freeze for easier review.
// AI Prompts: Show C# code using Aspose.Cells to set PageSetup.PrintArea, PageSetup.PrintTitleRows, and FreezePanes so the header row repeats on print and stays frozen. | Provide a step‑by‑step guide to align the worksheet view with the printed output by freezing rows after defining a print area in Aspose.Cells. | Explain how to configure print titles and freeze panes together for a printable Excel report in Aspose.Cells for .NET.

using System;
using Aspose.Cells;

// Learn how to set a print area (A1:B20), repeat the first row on every printed page, and freeze that header row in an Aspose.Cells workbook using C#.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add header row
        worksheet.Cells["A1"].PutValue("Header1");
        worksheet.Cells["B1"].PutValue("Header2");

        // Add sample data rows (rows 2‑20)
        for (int i = 2; i <= 20; i++)
        {
            worksheet.Cells[$"A{i}"].PutValue($"Data{i - 1}A");
            worksheet.Cells[$"B{i}"].PutValue($"Data{i - 1}B");
        }

        // Set the print area to include the header and all data rows
        worksheet.PageSetup.PrintArea = "A1:B20";

        // Ensure the header row repeats on each printed page
        worksheet.PageSetup.PrintTitleRows = "$1:$1";

        // Freeze the header row so the view aligns with the printed output
        // Freeze at cell A2, freezing 1 row (the header) and 0 columns
        worksheet.FreezePanes("A2", 1, 0);

        // Save the workbook
        workbook.Save("PrintAreaAndFreezeDemo.xlsx");
    }
}
