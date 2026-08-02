// Title: C# – Add a Calculated Field to an Aspose.Cells Pivot Table to Categorize Sales (High, Medium, Low)
// Description: This example creates a workbook, fills a small sales dataset, builds a pivot table on A1:C5, adds Region and Product as row/column fields, includes Sales as a data field, and defines a calculated field named **SalesCategory** using a nested IF formula ("=IF(Sales>1000,\"High\",IF(Sales>500,\"Medium\",\"Low\"))"). The pivot is refreshed, calculated, and saved as an Excel file.
// Keywords: Aspose.Cells calculated field | pivot table IF formula C# | AddCalculatedField Aspose.Cells | sales categorization pivot | .NET Excel pivot table | nested IF Excel formula | regional sales report C# | Excel dashboard Aspose.Cells
// Common Searches: how to add a calculated field in Aspose.Cells pivot table | C# Aspose.Cells IF formula for sales categories | Aspose.Cells AddCalculatedField example | pivot table sales tier classification .NET | create sales category field in Excel using Aspose
// Developer Intent: Generate a pivot table and attach a calculated field that classifies each sales value as High, Medium, or Low.
// Use Cases: Produce a regional sales summary where each amount is tagged with a performance tier for quick decision‑making. | Build an interactive Excel dashboard that shows product‑region sales alongside categorical labels for conditional formatting. | Export a workbook that can be consumed by BI tools, with a pre‑calculated sales‑category field for downstream analysis.
// AI Prompts: Write C# code with Aspose.Cells to add a calculated field called 'SalesCategory' to a pivot table using a nested IF formula that returns High, Medium, or Low. | Explain the steps required to refresh and recalculate a pivot table after adding a calculated field with Aspose.Cells. | Show how to adjust the threshold values in the IF formula to create custom sales categories in an Aspose.Cells pivot table.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotCalculatedFieldDemo
{
    // This example creates a workbook, fills a small sales dataset, builds a pivot table on A1:C5, adds Region and Product as row/column fields, includes Sales as a data field, and defines a calculated field named **SalesCategory** using a nested IF formula ("=IF(Sales>1000,\"High\",IF(Sales>500,\"Medium\",\"Low\"))"). The pivot is refreshed, calculated, and saved as an Excel file.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data for the pivot table
            // Header row
            cells["A1"].Value = "Region";
            cells["B1"].Value = "Product";
            cells["C1"].Value = "Sales";

            // Data rows
            cells["A2"].Value = "North";
            cells["B2"].Value = "Widget";
            cells["C2"].Value = 1200;

            cells["A3"].Value = "North";
            cells["B3"].Value = "Gadget";
            cells["C3"].Value = 800;

            cells["A4"].Value = "South";
            cells["B4"].Value = "Widget";
            cells["C4"].Value = 450;

            cells["A5"].Value = "South";
            cells["B5"].Value = "Gadget";
            cells["C5"].Value = 1500;

            // Create a pivot table based on the data range
            // Place the pivot table starting at cell E3
            int pivotIndex = sheet.PivotTables.Add("A1:C5", "E3", "SalesPivot");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Add fields to the pivot table
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Region");
            pivotTable.AddFieldToArea(PivotFieldType.Column, "Product");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

            // Add a calculated field that categorizes sales:
            // High  : Sales > 1000
            // Medium: 500 < Sales <= 1000
            // Low   : Sales <= 500
            // The IF function in Excel syntax is used.
            string formula = "=IF(Sales>1000,\"High\",IF(Sales>500,\"Medium\",\"Low\"))";
            pivotTable.AddCalculatedField("SalesCategory", formula, true);

            // Refresh and calculate the pivot table to apply the new field
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook
            workbook.Save("PivotTable_With_CalculatedField.xlsx");
        }
    }
}
