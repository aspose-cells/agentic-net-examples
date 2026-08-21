// Title: Aspose.Cells for .NET – Display PivotTable Data Fields as Percent of Total
// Description: C# sample that builds a workbook, inserts sample rows, creates a PivotTable and configures each data field to use the PercentageOfTotal display mode via ShowValuesSetting. The pivot is refreshed, calculated, and saved with values normalized as percentages.
// Keywords: Aspose.Cells | PivotTable | ShowValuesSetting | PercentageOfTotal | C# | .NET | data field percentage | normalize pivot values | Excel automation | pivot calculation type
// Common Searches: Aspose.Cells set pivot data field to percentage | C# show pivot values as percent of total | How to use ShowValuesSetting in Aspose.Cells | Normalize pivot table columns Aspose.Cells .NET | PivotFieldDataDisplayFormat PercentageOfTotal example
// Developer Intent: Apply a percent‑of‑overall‑total format to every data field in a PivotTable using Aspose.Cells.
// Use Cases: Sales analysis where each amount is shown as its share of the total revenue. | Budget reports that compare department expenditures as normalized percentages. | Dashboard visualizations that highlight product contribution ratios across categories.
// AI Prompts: Generate C# code with Aspose.Cells that sets ShowValuesSetting.CalculationType to PercentageOfTotal for all PivotTable data fields. | Explain the impact of ShowValuesSetting on pivot calculations and the required steps to refresh the pivot after modification. | Provide a concise example that iterates over PivotTable.DataFields, applies a percentage display format, and saves the workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// C# sample that builds a workbook, inserts sample rows, creates a PivotTable and configures each data field to use the PercentageOfTotal display mode via ShowValuesSetting. The pivot is refreshed, calculated, and saved with values normalized as percentages.
class SetPivotShowValuesAsPercent
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the pivot table
        // Columns: Category, SubCategory, Amount
        sheet.Cells["A1"].Value = "Category";
        sheet.Cells["B1"].Value = "SubCategory";
        sheet.Cells["C1"].Value = "Amount";

        sheet.Cells["A2"].Value = "Fruit";
        sheet.Cells["B2"].Value = "Apple";
        sheet.Cells["C2"].Value = 120;

        sheet.Cells["A3"].Value = "Fruit";
        sheet.Cells["B3"].Value = "Orange";
        sheet.Cells["C3"].Value = 150;

        sheet.Cells["A4"].Value = "Vegetable";
        sheet.Cells["B4"].Value = "Carrot";
        sheet.Cells["C4"].Value = 80;

        sheet.Cells["A5"].Value = "Vegetable";
        sheet.Cells["B5"].Value = "Broccoli";
        sheet.Cells["C5"].Value = 90;

        // Add a pivot table based on the data range
        PivotTableCollection pivots = sheet.PivotTables;
        int pivotIndex = pivots.Add("A1:C5", "E3", "PivotTable1");
        PivotTable pivot = pivots[pivotIndex];

        // Add fields to the pivot table
        pivot.AddFieldToArea(PivotFieldType.Row, "Category");          // Row field
        pivot.AddFieldToArea(PivotFieldType.Column, "SubCategory");   // Column field
        pivot.AddFieldToArea(PivotFieldType.Data, "Amount");          // Data field

        // Set ShowValuesAsPercent (percentage of total) for every data field
        foreach (PivotField dataField in pivot.DataFields)
        {
            // Use the ShowValuesSetting property to define the calculation type
            dataField.ShowValuesSetting.CalculationType = PivotFieldDataDisplayFormat.PercentageOfTotal;
        }

        // Refresh and calculate the pivot table to apply changes
        pivot.RefreshData();
        pivot.CalculateData();

        // Save the workbook
        workbook.Save("PivotTable_ShowValuesAsPercent.xlsx");
    }
}
