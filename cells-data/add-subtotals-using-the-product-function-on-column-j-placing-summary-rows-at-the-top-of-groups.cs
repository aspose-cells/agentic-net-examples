// Title: Add product subtotals for column J with summary rows above each group using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that groups rows by column A and inserts a Product subtotal for column J, placing the subtotal row before each group with Aspose.Cells. | Show how to call Cells.Subtotal in Aspose.Cells for .NET to calculate the product of values in column J and set summaryBelowData to false. | Create a workbook that writes sample data, applies product subtotals at the top of each group, and saves the file as Subtotal_Product_Top.xlsx using Aspose.Cells.
// Common Searches: aspnet cells subtotal product function column J top of group | c# Aspose.Cells how to place subtotal rows above groups | using Cells.Subtotal to calculate product totals in Excel with .NET | Aspose.Cells example grouping by column A and adding product subtotals | summaryBelowData false Aspose.Cells subtotal example
// Tags: Aspose.Cells Cells.Subtotal product aggregation | C# Excel product subtotal top of group | grouped subtotal rows Aspose.Cells | summaryBelowData false usage | Excel workbook product subtotal .NET

using System;
using Aspose.Cells;

namespace AsposeCellsSubtotalDemo
{
    // The example creates a new workbook, fills column A with group names and column J with numeric values, defines the data range, and uses Cells.Subtotal to group by column A, compute the Product of column J, place the subtotal rows above each group (summaryBelowData = false), then saves the result as Subtotal_Product_Top.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Sample data: column A will be the grouping field,
            // column J (index 9) will contain the values to subtotal.
            // Header row
            cells["A1"].PutValue("Group");
            cells["J1"].PutValue("Amount");

            // Data rows
            object[,] data = new object[,]
            {
                { "North", 2 },
                { "North", 3 },
                { "South", 4 },
                { "South", 5 },
                { "East",  6 },
                { "East",  7 }
            };

            // Populate the worksheet (rows start at index 1 because row 0 is header)
            for (int i = 0; i < data.GetLength(0); i++)
            {
                // Group column (A)
                cells[i + 1, 0].PutValue(data[i, 0]);
                // Value column (J)
                cells[i + 1, 9].PutValue(data[i, 1]);
            }

            // Define the cell area that contains the data (including headers)
            // StartRow = 0, StartColumn = 0 (A), EndRow = last data row, EndColumn = 9 (J)
            int lastRow = data.GetLength(0); // number of data rows
            CellArea area = new CellArea
            {
                StartRow = 0,
                StartColumn = 0,
                EndRow = lastRow,
                EndColumn = 9
            };

            // Apply subtotals:
            // - Group by column 0 (Group)
            // - Use Product function on column J (index 9)
            // - Do not replace existing subtotals, no page breaks,
            // - Place summary rows at the top of each group (summaryBelowData = false)
            cells.Subtotal(
                area,
                groupBy: 0,
                function: ConsolidationFunction.Product,
                totalList: new int[] { 9 },
                replace: false,
                pageBreaks: false,
                summaryBelowData: false);

            // Save the workbook
            workbook.Save("Subtotal_Product_Top.xlsx");
        }
    }
}
