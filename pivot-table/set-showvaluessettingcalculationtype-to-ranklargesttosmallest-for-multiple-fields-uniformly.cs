// Title: C# Aspose.Cells: Apply RankLargestToSmallest to All Pivot Table Data Fields
// Description: Shows how to create a workbook, add a pivot table with multiple data fields, and use a loop to set ShowValuesSetting.CalculationType to RankLargestToSmallest for every field, then refresh, calculate, and save the Excel file.
// Keywords: Aspose.Cells | C# | .NET | PivotTable | ShowValuesSetting | CalculationType | RankLargestToSmallest | multiple data fields | pivot ranking descending | Excel report
// Common Searches: Aspose.Cells set RankLargestToSmallest for all pivot fields | C# loop to apply ShowValuesSetting.CalculationType in pivot table | rank pivot table values descending Aspose.Cells | apply same ranking to multiple data fields in Excel pivot | Aspose.Cells .NET pivot table ranking example
// Developer Intent: Apply the RankLargestToSmallest calculation to every data field in a pivot table with a single operation.
// Use Cases: Generate an Excel report where all pivot data fields are automatically ranked from highest to lowest. | Simplify code by configuring ranking for multiple fields in one loop instead of individual settings. | Create dashboards that need consistent descending rankings across sum, count, average, or custom calculations.
// AI Prompts: Write C# code using Aspose.Cells to set ShowValuesSetting.CalculationType = RankLargestToSmallest for all pivot data fields. | Provide an example that adds several data fields to a pivot table and applies the same ranking calculation to each via a loop. | Explain the steps to refresh and recalculate a pivot table after changing ShowValuesSetting in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // Shows how to create a workbook, add a pivot table with multiple data fields, and use a loop to set ShowValuesSetting.CalculationType to RankLargestToSmallest for every field, then refresh, calculate, and save the Excel file.
    public class SetRankLargestToSmallestForMultipleFields
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

            // Populate sample data
            // Columns: Category, SubCategory, Amount
            sheet.Cells["A1"].Value = "Category";
            sheet.Cells["B1"].Value = "SubCategory";
            sheet.Cells["C1"].Value = "Amount";

            sheet.Cells["A2"].Value = "Food";
            sheet.Cells["B2"].Value = "Fruit";
            sheet.Cells["C2"].Value = 120;

            sheet.Cells["A3"].Value = "Food";
            sheet.Cells["B3"].Value = "Vegetable";
            sheet.Cells["C3"].Value = 80;

            sheet.Cells["A4"].Value = "Beverage";
            sheet.Cells["B4"].Value = "Tea";
            sheet.Cells["C4"].Value = 150;

            sheet.Cells["A5"].Value = "Beverage";
            sheet.Cells["B5"].Value = "Coffee";
            sheet.Cells["C5"].Value = 200;

            sheet.Cells["A6"].Value = "Food";
            sheet.Cells["B6"].Value = "Meat";
            sheet.Cells["C6"].Value = 300;

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:C6", "E3", "PivotTable1");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Add fields to the pivot table
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
            pivotTable.AddFieldToArea(PivotFieldType.Column, "SubCategory");

            // Add first data field (Sum)
            int dataFieldIdx1 = pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");
            pivotTable.DataFields[dataFieldIdx1].DisplayName = "Amount Sum";

            // Add second data field (Count)
            int dataFieldIdx2 = pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");
            pivotTable.DataFields[dataFieldIdx2].DisplayName = "Amount Count";
            pivotTable.DataFields[dataFieldIdx2].Function = ConsolidationFunction.Count;

            // Set RankLargestToSmallest for all data fields
            foreach (PivotField dataField in pivotTable.DataFields)
            {
                dataField.ShowValuesSetting.CalculationType = PivotFieldDataDisplayFormat.RankLargestToSmallest;
            }

            // Refresh and calculate the pivot table data
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook
            string outputPath = "PivotMultipleFields_RankLargestToSmallest.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
    }
}
