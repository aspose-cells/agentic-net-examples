// Title: Aspose.Cells C# – Add IF‑Based Calculated Field to Pivot Table for Sales Category (High, Medium, Low)
// Description: Demonstrates how to create a workbook with region‑sales data, build a pivot table, and programmatically add a calculated field named **SalesCategory** using the IF formula `=IF(Sales>200,"High",IF(Sales>100,"Medium","Low"))`. The example shows adding the field to the data area, refreshing the pivot, recalculating, and saving the file as an Excel workbook.
// Keywords: Aspose.Cells calculated field | C# pivot table IF formula | sales category pivot | add calculated field Aspose.Cells | .NET Excel pivot table example | IF function in pivot table | categorize sales high medium low
// Common Searches: Aspose.Cells add calculated field with IF | C# pivot table sales category example | how to use IF formula in Aspose.Cells pivot | refresh pivot table after adding calculated field Aspose.Cells | programmatically create pivot table in .NET
// Developer Intent: Create a pivot table in Aspose.Cells and attach an IF‑based calculated field that classifies each sales value as High, Medium, or Low.
// Use Cases: Generate regional sales reports where analysts can filter by High, Medium, or Low sales categories. | Automate workbook production for multiple datasets, applying the same SalesCategory field to each pivot table. | Export Excel workbooks with pre‑calculated sales tiers for downstream BI tools.
// AI Prompts: Write C# code using Aspose.Cells to add a calculated field called 'SalesCategory' with the formula IF(Sales>200,"High",IF(Sales>100,"Medium","Low")) to an existing pivot table and refresh it. | Show how to change the threshold values in the SalesCategory IF formula for different business rules in Aspose.Cells. | Provide a step‑by‑step guide to read back the values of the 'SalesCategory' calculated field after the pivot table is calculated.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook with region‑sales data, build a pivot table, and programmatically add a calculated field named **SalesCategory** using the IF formula `=IF(Sales>200,"High",IF(Sales>100,"Medium","Low"))`. The example shows adding the field to the data area, refreshing the pivot, recalculating, and saving the file as an Excel workbook.
    public class PivotTableCalculatedFieldDemo
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data for the pivot table
            cells["A1"].Value = "Region";
            cells["B1"].Value = "Sales";

            cells["A2"].Value = "North";
            cells["B2"].Value = 250;   // High
            cells["A3"].Value = "North";
            cells["B3"].Value = 150;   // Medium
            cells["A4"].Value = "South";
            cells["B4"].Value = 80;    // Low
            cells["A5"].Value = "South";
            cells["B5"].Value = 300;   // High
            cells["A6"].Value = "East";
            cells["B6"].Value = 120;   // Medium
            cells["A7"].Value = "West";
            cells["B7"].Value = 60;    // Low

            // Add a pivot table based on the data range A1:B7, place it at E3
            int pivotIndex = sheet.PivotTables.Add("A1:B7", "E3", "SalesPivot");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Add fields to the pivot table
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Region");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

            // Add a calculated field that categorizes sales using IF
            // Formula: IF(Sales>200,"High",IF(Sales>100,"Medium","Low"))
            string calcFieldName = "SalesCategory";
            string calcFormula = "=IF(Sales>200,\"High\",IF(Sales>100,\"Medium\",\"Low\"))";

            // Drag the calculated field to the data area immediately
            pivotTable.AddCalculatedField(calcFieldName, calcFormula, true);

            // Refresh the pivot table data to apply the new calculated field
            pivotTable.RefreshData();

            // Recalculate the pivot table to apply changes
            pivotTable.CalculateData();

            // Save the workbook
            workbook.Save("PivotTableWithSalesCategory.xlsx");
        }
    }
}
