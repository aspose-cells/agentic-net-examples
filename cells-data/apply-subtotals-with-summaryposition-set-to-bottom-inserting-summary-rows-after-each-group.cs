// Title: Create Excel subtotals with summary rows positioned at the bottom of each category group using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that groups rows by a column, calculates a SUM subtotal on another column, and inserts the subtotal rows after each group (summary below data). | Show how to retrieve the SubtotalSetting after applying subtotals and display its SummaryBelowData property to verify bottom placement.
// Common Searches: how to add subtotal rows below each group in an Excel file using Aspose.Cells C# | Aspose.Cells C# subtotal summary position bottom example | group data by column and insert sum subtotal at the end of each group with Aspose.Cells for .NET
// Tags: Aspose.Cells subtotal bottom placement C# | C# apply sum subtotal by column Aspose.Cells | Excel worksheet group subtotals Aspose.Cells | retrieve SubtotalSetting Aspose.Cells C#

using System;
using Aspose.Cells;

// The example creates a workbook, fills it with Category, Item, and Amount data, defines the range A1:C5, and uses Cells.Subtotal to group rows by the Category column, apply a SUM subtotal on the Amount column, and place the summary rows below each group. It then retrieves the SubtotalSetting to confirm the SummaryBelowData flag is true and saves the file as SubtotalBottomDemo.xlsx.
public class SubtotalBottomDemo
{
    public static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Add header row
        cells["A1"].PutValue("Category");
        cells["B1"].PutValue("Item");
        cells["C1"].PutValue("Amount");

        // Sample data
        object[,] data = new object[,]
        {
            { "North", "Widget", 5000 },
            { "North", "Gadget", 3000 },
            { "South", "Widget", 6000 },
            { "South", "Gadget", 4000 },
            { "West",  "Widget", 4500 }
        };

        // Populate data starting from row 2 (zero‑based index 1)
        for (int i = 0; i < data.GetLength(0); i++)
        {
            cells[i + 1, 0].PutValue(data[i, 0]);
            cells[i + 1, 1].PutValue(data[i, 1]);
            cells[i + 1, 2].PutValue(data[i, 2]);
        }

        // Define the cell area that includes the header and data (A1:C5)
        CellArea area = CellArea.CreateCellArea("A1", "C5");

        // Apply subtotals:
        // - Group by column 0 (Category)
        // - Use SUM function
        // - Subtotal on column 2 (Amount)
        // - Do not replace existing subtotals
        // - No page breaks between groups
        // - Summary rows placed below data (bottom)
        cells.Subtotal(area, 0, ConsolidationFunction.Sum, new int[] { 2 }, false, false, true);

        // Retrieve and display the subtotal setting to confirm the summary position
        SubtotalSetting setting = cells.RetrieveSubtotalSetting(area);
        Console.WriteLine("SummaryBelowData (bottom): " + setting.SummaryBelowData);

        // Save the workbook
        workbook.Save("SubtotalBottomDemo.xlsx");
    }
}
