// Title: Aspose.Cells C# PivotTable – Show Values as Percentage of Column Total
// Description: Creates a workbook with sample sales data, builds a PivotTable (Product rows, Quarter columns, Sales values) and uses PivotField.ShowValuesAs with the PercentageOfColumn format to display each sales figure as a share of its column total, then refreshes, calculates, and saves the file.
// Keywords: Aspose.Cells PivotTable C# | ShowValuesAs PercentageOfColumn | pivot table column percentage | Aspose.Cells display values as percent | .NET Excel pivot percentages | RefreshData CalculateData Aspose.Cells
// Common Searches: Aspose.Cells show pivot values as percent of column | C# PivotField.ShowValuesAs PercentageOfColumn example | How to set column total percentage in Aspose.Cells pivot | Refresh and calculate pivot after changing display format Aspose.Cells | Aspose.Cells pivot table percentage of column total
// Developer Intent: Configure a PivotTable data field to display values as a percentage of the column total.
// Use Cases: Sales analysis where each product's revenue is shown as its share of quarterly totals. | Financial reporting that presents expense categories as a percentage of monthly columns. | Dashboard exports that require pivot percentages instead of raw numbers.
// AI Prompts: Generate C# code with Aspose.Cells to set a pivot table data field to PercentageOfColumn and refresh the table. | Explain the parameters of PivotField.ShowValuesAs for calculating column‑total percentages. | Show how to format the resulting percentages to two decimal places in the exported Excel workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotDemo
{
    // Creates a workbook with sample sales data, builds a PivotTable (Product rows, Quarter columns, Sales values) and uses PivotField.ShowValuesAs with the PercentageOfColumn format to display each sales figure as a share of its column total, then refreshes, calculates, and saves the file.
    class ShowValuesAsPercentageOfColumn
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data for the pivot table
            // Columns: Product, Quarter, Sales
            cells["A1"].Value = "Product";
            cells["B1"].Value = "Quarter";
            cells["C1"].Value = "Sales";

            cells["A2"].Value = "Laptop";
            cells["B2"].Value = "Q1";
            cells["C2"].Value = 1200;

            cells["A3"].Value = "Laptop";
            cells["B3"].Value = "Q2";
            cells["C3"].Value = 1500;

            cells["A4"].Value = "Phone";
            cells["B4"].Value = "Q1";
            cells["C4"].Value = 800;

            cells["A5"].Value = "Phone";
            cells["B5"].Value = "Q2";
            cells["C5"].Value = 950;

            cells["A6"].Value = "Tablet";
            cells["B6"].Value = "Q1";
            cells["C6"].Value = 600;

            cells["A7"].Value = "Tablet";
            cells["B7"].Value = "Q2";
            cells["C7"].Value = 700;

            // Add a pivot table to the worksheet
            // Source range: A1:C7, Destination: E3, Name: SalesPivot
            int pivotIndex = sheet.PivotTables.Add("A1:C7", "E3", "SalesPivot");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Add fields to the pivot table
            // Row field: Product
            int rowFieldIndex = pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
            // Column field: Quarter
            int columnFieldIndex = pivotTable.AddFieldToArea(PivotFieldType.Column, "Quarter");
            // Data field: Sales
            int dataFieldIndex = pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

            // Retrieve the data field object
            PivotField dataField = pivotTable.DataFields[dataFieldIndex];

            // Configure the data field to show values as percentage of column total
            // Using the ShowValuesAs method as required
            dataField.ShowValuesAs(
                PivotFieldDataDisplayFormat.PercentageOfColumn, // display format
                columnFieldIndex,                               // base field (the column field)
                PivotItemPositionType.Next,                    // base item position type (default)
                0);                                             // base item index (not used for PercentageOfColumn)

            // Refresh the pivot table data and calculate the results
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook to a file
            workbook.Save("Pivot_ShowValuesAs_PercentageOfColumn.xlsx");
        }
    }
}
