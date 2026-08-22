// Title: Create a pivot table and add a nested IF calculated field to classify sales as High, Medium, or Low with Aspose.Cells in C#
// AI Prompts: Generate C# code that builds a workbook, inserts product and sales data, creates a pivot table, and defines a calculated field named SalesCategory using the formula =IF(Sales>2000,"High",IF(Sales>1000,"Medium","Low")) with Aspose.Cells. | Show how to programmatically refresh the pivot table data and recalculate its values after adding the calculated field using the Aspose.Cells API. | Provide the steps to save the workbook containing the pivot table and the SalesCategory calculated field to an .xlsx file.
// Common Searches: aspocells c# add calculated field with nested if to pivot table | how to label sales high medium low in an Aspose.Cells pivot table | example of using IF function in Aspose.Cells pivot table calculated field | c# Aspose.Cells pivot table custom field for sales categorization
// Tags: insert IF expression into Aspose.Cells pivot | sales level classification pivot table C# | pivot table formula syntax Aspose.Cells | export workbook with pivot Aspose.Cells | place custom field in pivot data area Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotCalculatedFieldDemo
{
    // // Demonstrates creating a workbook, populating product and sales data, building a pivot table, adding a nested IF calculated field named SalesCategory to classify sales as High, Medium, or Low, refreshing and recalculating the pivot, and saving the workbook as an .xlsx file.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data
            cells["A1"].Value = "Product";
            cells["B1"].Value = "Sales";

            cells["A2"].Value = "Apple";
            cells["B2"].Value = 2500;

            cells["A3"].Value = "Banana";
            cells["B3"].Value = 1500;

            cells["A4"].Value = "Cherry";
            cells["B4"].Value = 800;

            // Create a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:B4", "D3", "SalesPivot");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Add fields to the pivot table
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

            // Add a calculated field that categorizes sales using IF
            // Formula: =IF(Sales>2000,"High",IF(Sales>1000,"Medium","Low"))
            string formula = "=IF(Sales>2000,\"High\",IF(Sales>1000,\"Medium\",\"Low\"))";
            pivotTable.AddCalculatedField("SalesCategory", formula, true);

            // Refresh and calculate the pivot table to apply changes
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook
            workbook.Save("PivotTable_With_SalesCategory.xlsx");
        }
    }
}
