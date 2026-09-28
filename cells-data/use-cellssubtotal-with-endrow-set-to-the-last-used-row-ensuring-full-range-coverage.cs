// Title: Use Aspose.Cells Subtotal in C# to sum the Sales column over a dynamically sized range by detecting the worksheet’s last used row
// AI Prompts: Write C# code that creates a workbook, populates it with data, uses MaxDataRow and MaxDataColumn to find the worksheet bounds, defines a CellArea covering the entire range, and calls Cells.Subtotal to group by the first column and sum the third column. | Show how to configure the Subtotal method to replace any existing subtotals, insert page breaks, and place the summary row below the data using Aspose.Cells. | Demonstrate saving the workbook as an .xlsx file after applying the subtotal operation.
// Common Searches: aspnet cells subtotal dynamic range last row c# | how to automatically set endRow for Cells.Subtotal in Aspose.Cells | c# Aspose.Cells sum subtotal by first column with page breaks | detect max data row and column before applying subtotal Aspose.Cells | apply subtotal to entire data set using CellArea in C#
// Tags: Aspose.Cells Cells.Subtotal dynamic range | C# worksheet bounds detection Aspose.Cells | subtotal sum by column Aspose.Cells | CellArea full data coverage | replace existing subtotals Aspose.Cells

using System;
using Aspose.Cells;

namespace AsposeCellsSubtotalDemo
{
    // The example creates a new workbook, fills it with Region, Product, and Sales data, determines the last occupied row and column via MaxDataRow/MaxDataColumn, defines a CellArea that spans the whole data set, and invokes Cells.Subtotal to group by the Region column and sum the Sales column. The call replaces any prior subtotals, adds page breaks, and places the summary below the data before saving the file as SubtotalFullRangeDemo.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate sample data (header + 5 rows, 3 columns)
            cells["A1"].PutValue("Region");
            cells["B1"].PutValue("Product");
            cells["C1"].PutValue("Sales");

            object[,] data = new object[,]
            {
                { "North", "Widget", 5000 },
                { "North", "Gadget", 3000 },
                { "South", "Widget", 6000 },
                { "South", "Gadget", 4000 },
                { "West",  "Widget", 4500 }
            };

            for (int r = 0; r < data.GetLength(0); r++)
            {
                for (int c = 0; c < data.GetLength(1); c++)
                {
                    cells[r + 1, c].PutValue(data[r, c]);
                }
            }

            // Determine the last used row and column
            int lastRow = cells.MaxDataRow;          // zero‑based index of the last row containing data
            int lastColumn = cells.MaxDataColumn;    // zero‑based index of the last column containing data

            // Define the cell area covering the entire data range
            CellArea area = new CellArea
            {
                StartRow = 0,
                StartColumn = 0,
                EndRow = lastRow,
                EndColumn = lastColumn
            };

            // Apply subtotal:
            // - Group by the first column (Region) -> groupBy = 0
            // - Use SUM function
            // - Subtotal the third column (Sales) -> totalList = {2}
            // - Replace existing subtotals, add page breaks, place summary below data
            cells.Subtotal(area, 0, ConsolidationFunction.Sum, new int[] { 2 }, true, true, true);

            // Save the workbook
            workbook.Save("SubtotalFullRangeDemo.xlsx");
        }
    }
}
