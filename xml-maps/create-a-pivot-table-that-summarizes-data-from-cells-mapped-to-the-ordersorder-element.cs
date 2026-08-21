// Title: Create a Pivot Table from XML‑Mapped /Orders/Order Data with Aspose.Cells (C#)
// Description: C# example that builds a workbook, adds an "Orders" sheet with OrderID, Customer, Region, and Amount columns, then creates a pivot table on a second sheet. The pivot groups orders by Region and sums Amount, uses a compact layout, refreshes the cache, and saves as OrdersPivotReport.xlsx.
// Keywords: Aspose.Cells pivot table C# | XML map to Excel Aspose | programmatic pivot table .NET | summarize order data by region | refresh calculate pivot Aspose.Cells | compact layout pivot Aspose | create workbook Aspose.Cells
// Common Searches: Aspose.Cells create pivot table from XML‑mapped data C# | How to add a pivot table with source range =Orders!A1:D5 using Aspose.Cells | Set row field and sum field in Aspose.Cells pivot table | Refresh and calculate pivot cache Aspose.Cells .NET | Compact layout for Aspose.Cells pivot table
// Developer Intent: Generate an Excel pivot table that summarizes XML‑mapped order records by Region and calculates total Amount using Aspose.Cells for .NET.
// Use Cases: Automated sales reporting that groups order totals by geographic region. | Server‑side generation of Excel dashboards from XML‑based order feeds. | Exporting summarized order metrics for downstream BI tools without manual Excel work.
// AI Prompts: Write C# code with Aspose.Cells to create a pivot table from a worksheet named 'Orders' (range A1:D5), placing 'Region' in rows and summing 'Amount'. | Explain how to refresh and calculate a pivot table after configuring fields in Aspose.Cells for .NET. | Show how to apply a compact layout to a pivot table created with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotExample
{
    // C# example that builds a workbook, adds an "Orders" sheet with OrderID, Customer, Region, and Amount columns, then creates a pivot table on a second sheet. The pivot groups orders by Region and sums Amount, uses a compact layout, refreshes the cache, and saves as OrdersPivotReport.xlsx.
    class Program
    {
        static void Main()
        {
            // 1. Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // 2. Add a worksheet that will hold the source XML‑mapped data.
            //    For demonstration we simulate the /Orders/Order element with a simple table.
            Worksheet dataSheet = workbook.Worksheets[0];
            dataSheet.Name = "Orders";

            // Header row (simulating fields from /Orders/Order)
            dataSheet.Cells["A1"].PutValue("OrderID");
            dataSheet.Cells["B1"].PutValue("Customer");
            dataSheet.Cells["C1"].PutValue("Region");
            dataSheet.Cells["D1"].PutValue("Amount");

            // Sample data rows
            dataSheet.Cells["A2"].PutValue(1001);
            dataSheet.Cells["B2"].PutValue("Alice");
            dataSheet.Cells["C2"].PutValue("North");
            dataSheet.Cells["D2"].PutValue(2500);

            dataSheet.Cells["A3"].PutValue(1002);
            dataSheet.Cells["B3"].PutValue("Bob");
            dataSheet.Cells["C3"].PutValue("South");
            dataSheet.Cells["D3"].PutValue(1800);

            dataSheet.Cells["A4"].PutValue(1003);
            dataSheet.Cells["B4"].PutValue("Charlie");
            dataSheet.Cells["C4"].PutValue("North");
            dataSheet.Cells["D4"].PutValue(3200);

            dataSheet.Cells["A5"].PutValue(1004);
            dataSheet.Cells["B5"].PutValue("Diana");
            dataSheet.Cells["C5"].PutValue("East");
            dataSheet.Cells["D5"].PutValue(1500);

            // 3. Add a new worksheet that will contain the pivot table.
            Worksheet pivotSheet = workbook.Worksheets.Add("PivotReport");

            // 4. Define the source range for the pivot table.
            //    The range includes the header row and all data rows.
            string sourceRange = $"=Orders!A1:D5";

            // 5. Add the pivot table to the pivot sheet.
            //    Parameters: source data, destination cell (top‑left of the pivot), pivot name.
            int pivotIndex = pivotSheet.PivotTables.Add(sourceRange, "A3", "OrdersPivot");
            PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

            // 6. Configure the pivot table fields:
            //    - Row area: Region (group orders by region)
            //    - Column area: Customer (optional, can be omitted)
            //    - Data area: Sum of Amount
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Region");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

            // 7. Optional: set a compact layout for better readability.
            pivotTable.ShowInCompactForm();

            // 8. Refresh the pivot cache and calculate the results.
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // 9. Save the workbook (lifecycle: save)
            workbook.Save("OrdersPivotReport.xlsx");
        }
    }
}
