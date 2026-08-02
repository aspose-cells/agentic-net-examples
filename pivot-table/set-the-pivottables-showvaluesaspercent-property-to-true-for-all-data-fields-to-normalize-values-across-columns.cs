// Title: Set ShowValuesAsPercent for All Data Fields in an Aspose.Cells PivotTable (C#)
// Description: Creates a workbook, adds sample data, builds a pivot table, and programmatically sets each data field's ShowValuesSetting to PercentageOfTotal. The pivot is refreshed, recalculated, and saved as an XLSX file.
// Keywords: Aspose.Cells PivotTable ShowValuesAsPercent | C# set pivot field percentage | Aspose.Cells PercentageOfTotal | .NET pivot table display as percent | normalize pivot values Aspose | ShowValuesSetting CalculationType
// Common Searches: Aspose.Cells set pivot values to percent of total | C# pivot table show values as percentage | How to enable ShowValuesAsPercent for all data fields in Aspose.Cells | PivotField ShowValuesSetting PercentageOfTotal example | Aspose.Cells pivot table percentage formatting
// Developer Intent: Apply a percentage‑of‑total display to every data field in a pivot table using Aspose.Cells for .NET.
// Use Cases: Transform monetary amounts into relative percentages for financial dashboards. | Standardize percentage formatting across multiple data fields when generating reports automatically. | Ensure the workbook reflects the new display format by refreshing and recalculating the pivot after changes.
// AI Prompts: Write C# code with Aspose.Cells that iterates over PivotTable.DataFields and sets ShowValuesSetting.CalculationType to PercentageOfTotal for each field. | Create a method that enables ShowValuesAsPercent for all data fields in an existing Aspose.Cells pivot table, then refreshes and calculates the pivot. | Explain step‑by‑step how to display pivot table values as a percentage of total using Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotDemo
{
    // Creates a workbook, adds sample data, builds a pivot table, and programmatically sets each data field's ShowValuesSetting to PercentageOfTotal. The pivot is refreshed, recalculated, and saved as an XLSX file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            Cells cells = sheet.Cells;
            cells["A1"].Value = "Category";
            cells["B1"].Value = "SubCategory";
            cells["C1"].Value = "Amount";

            cells["A2"].Value = "Food";
            cells["B2"].Value = "Fruit";
            cells["C2"].Value = 1200;

            cells["A3"].Value = "Food";
            cells["B3"].Value = "Vegetable";
            cells["C3"].Value = 800;

            cells["A4"].Value = "Beverage";
            cells["B4"].Value = "Tea";
            cells["C4"].Value = 500;

            cells["A5"].Value = "Beverage";
            cells["B5"].Value = "Coffee";
            cells["C5"].Value = 700;

            // Add a pivot table based on the data range
            PivotTableCollection pivots = sheet.PivotTables;
            int pivotIndex = pivots.Add("A1:C5", "E3", "PivotTable1");
            PivotTable pivotTable = pivots[pivotIndex];

            // Add fields to the pivot table
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
            pivotTable.AddFieldToArea(PivotFieldType.Column, "SubCategory");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

            // Iterate over all data fields and set ShowValuesSetting to display values as percentage of total
            foreach (PivotField dataField in pivotTable.DataFields)
            {
                dataField.ShowValuesSetting.CalculationType = PivotFieldDataDisplayFormat.PercentageOfTotal;
            }

            // Refresh and calculate the pivot table to apply the settings
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook
            workbook.Save("PivotTable_ShowValuesAsPercent.xlsx");
        }
    }
}
