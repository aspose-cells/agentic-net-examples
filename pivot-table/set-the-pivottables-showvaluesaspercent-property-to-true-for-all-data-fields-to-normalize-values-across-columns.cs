// Title: Set all PivotTable data fields to display values as percentages of total using Aspose.Cells for .NET (C#)
// AI Prompts: Create a workbook, add sample data, generate a pivot table, and configure each data field to show values as PercentageOfTotal with Aspose.Cells in C#. | Iterate over the PivotTable.DataFields collection and assign ShowValuesSetting.CalculationType = PivotFieldDataDisplayFormat.PercentageOfTotal, then refresh and calculate the pivot. | Save the workbook after applying the percentage display setting to the pivot table.
// Common Searches: Aspose.Cells C# how to set pivot table data fields to percentage of total | C# example for showing pivot table values as percent using Aspose.Cells | Set ShowValuesSetting.CalculationType to PercentageOfTotal for all pivot data fields Aspose | Normalize pivot table values across columns with Aspose.Cells .NET
// Tags: Aspose.Cells pivot table percentage display | C# ShowValuesSetting CalculationType PercentageOfTotal | Aspose.Cells refresh calculate pivot | Excel workbook save Aspose.Cells C# | normalize pivot values across columns Aspose

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The sample creates a workbook, fills it with sales data, adds a pivot table, assigns row, column, and data fields, sets each data field's ShowValuesSetting.CalculationType to PercentageOfTotal, refreshes and calculates the pivot, and saves the result as PivotTableShowValuesAsPercent.xlsx.
class PivotTableShowValuesAsPercentDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the pivot table
        Cells cells = sheet.Cells;
        cells["A1"].Value = "Category";
        cells["B1"].Value = "Region";
        cells["C1"].Value = "Sales";

        cells["A2"].Value = "Electronics"; cells["B2"].Value = "North"; cells["C2"].Value = 1200;
        cells["A3"].Value = "Electronics"; cells["B3"].Value = "South"; cells["C3"].Value = 1500;
        cells["A4"].Value = "Furniture";   cells["B4"].Value = "North"; cells["C4"].Value = 800;
        cells["A5"].Value = "Furniture";   cells["B5"].Value = "South"; cells["C5"].Value = 950;

        // Add a pivot table based on the data range
        int pivotIndex = sheet.PivotTables.Add("A1:C5", "E3", "SalesPivot");
        PivotTable pivotTable = sheet.PivotTables[pivotIndex];

        // Add fields to the pivot table
        pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
        pivotTable.AddFieldToArea(PivotFieldType.Column, "Region");
        pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

        // Iterate over all data fields and set ShowValuesSetting to display values as percentage of total
        foreach (PivotField dataField in pivotTable.DataFields)
        {
            dataField.ShowValuesSetting.CalculationType = PivotFieldDataDisplayFormat.PercentageOfTotal;
        }

        // Refresh and calculate the pivot table to apply the changes
        pivotTable.RefreshData();
        pivotTable.CalculateData();

        // Save the workbook
        workbook.Save("PivotTableShowValuesAsPercent.xlsx");
    }
}
