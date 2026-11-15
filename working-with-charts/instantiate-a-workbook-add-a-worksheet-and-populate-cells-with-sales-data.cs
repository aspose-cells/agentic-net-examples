// Title: Generate an Excel workbook with a "Sales" worksheet and insert tabular sales data using Aspose.Cells for .NET (C#)
// Common Searches: how to create a new workbook and add a named worksheet with Aspose.Cells in C# | populate Excel sheet from a two‑dimensional object array using Aspose.Cells .NET | write header row and sales rows to an .xlsx file with Aspose.Cells C# example | save workbook as SalesData.xlsx programmatically with Aspose.Cells | add worksheet, set cell values, and export to Excel using Aspose.Cells for .NET
// Tags: create workbook and add named worksheet Aspose.Cells | write header row to cells using PutValue C# | populate worksheet from two‑dimensional array Aspose.Cells | export workbook to .xlsx file Aspose.Cells | sample sales data insertion Aspose.Cells C#

using Aspose.Cells;
using System;

// The code creates a new Workbook, adds a worksheet called "Sales", writes a header row, fills the sheet with a predefined two‑dimensional array of product, region, units sold, and revenue data, and saves the result as "SalesData.xlsx" using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add a new worksheet and name it "Sales"
        int sheetIndex = workbook.Worksheets.Add();
        Worksheet sheet = workbook.Worksheets[sheetIndex];
        sheet.Name = "Sales";

        // Write header row
        sheet.Cells["A1"].PutValue("Product");
        sheet.Cells["B1"].PutValue("Region");
        sheet.Cells["C1"].PutValue("Units Sold");
        sheet.Cells["D1"].PutValue("Revenue");

        // Sample sales data
        object[,] salesData = new object[,]
        {
            { "Widget", "North", 120, 2400 },
            { "Widget", "South", 85, 1700 },
            { "Gadget", "North", 150, 4500 },
            { "Gadget", "South", 130, 3900 }
        };

        // Populate cells with sales data starting from row 2 (index 1)
        for (int row = 0; row < salesData.GetLength(0); row++)
        {
            for (int col = 0; col < salesData.GetLength(1); col++)
            {
                sheet.Cells[row + 1, col].PutValue(salesData[row, col]);
            }
        }

        // Save the workbook to a file
        workbook.Save("SalesData.xlsx");
    }
}
