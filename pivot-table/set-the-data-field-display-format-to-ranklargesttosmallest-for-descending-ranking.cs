// Title: How to set a pivot table data field to RankLargestToSmallest (descending rank) using Aspose.Cells for .NET (C#)
// AI Prompts: Create a new workbook, add a pivot table, and configure its data field to use the RankLargestToSmallest calculation type with Aspose.Cells in C#. | Update an existing Aspose.Cells pivot table so that the Quantity data field displays ranks from highest to lowest. | Generate an Excel file named PivotRankLargestToSmallest.xlsx where the pivot table ranks items by descending quantity using the Aspose.Cells .NET API.
// Common Searches: Aspose.Cells C# set pivot table data field ranking to descending | How to use RankLargestToSmallest with Aspose.Cells pivot tables in .NET | C# example ranking pivot table values from largest to smallest using Aspose.Cells | ShowValuesSetting CalculationType RankLargestToSmallest Aspose.Cells example | Create Excel pivot table with descending rank display using Aspose.Cells for .NET
// Tags: Aspose.Cells pivot table ranking | C# set PivotField ShowValuesSetting CalculationType | RankLargestToSmallest Aspose.Cells | descending rank Excel pivot .NET | pivot table data field ranking configuration

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotRankExample
{
    // The C# program creates a workbook, fills it with sample Category, Item, and Quantity data, adds a pivot table, assigns Item to rows, Category to columns, and Quantity to the data area, then sets the data field's ShowValuesSetting.CalculationType to RankLargestToSmallest so the pivot displays descending ranks, refreshes and calculates the pivot, and saves the result as PivotRankLargestToSmallest.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data for the pivot table
            cells["A1"].Value = "Category";
            cells["B1"].Value = "Item";
            cells["C1"].Value = "Quantity";

            cells["A2"].Value = "Fruit";
            cells["B2"].Value = "Apple";
            cells["C2"].Value = 30;

            cells["A3"].Value = "Fruit";
            cells["B3"].Value = "Banana";
            cells["C3"].Value = 20;

            cells["A4"].Value = "Fruit";
            cells["B4"].Value = "Orange";
            cells["C4"].Value = 50;

            cells["A5"].Value = "Vegetable";
            cells["B5"].Value = "Carrot";
            cells["C5"].Value = 40;

            cells["A6"].Value = "Vegetable";
            cells["B6"].Value = "Broccoli";
            cells["C6"].Value = 10;

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:C6", "E3", "PivotTable1");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Add fields to the pivot table
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Item");          // Row field
            pivotTable.AddFieldToArea(PivotFieldType.Column, "Category"); // Column field
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Quantity");   // Data field

            // Retrieve the data field (the first data field)
            PivotField dataField = pivotTable.DataFields[0];

            // Set the data field to display ranking from largest to smallest
            dataField.ShowValuesSetting.CalculationType = PivotFieldDataDisplayFormat.RankLargestToSmallest;

            // Refresh and calculate the pivot table data
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook to a file
            workbook.Save("PivotRankLargestToSmallest.xlsx");
        }
    }
}
