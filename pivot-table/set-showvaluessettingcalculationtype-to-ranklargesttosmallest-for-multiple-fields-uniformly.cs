// Title: How to set ShowValuesSetting.CalculationType to RankLargestToSmallest for multiple pivot data fields using Aspose.Cells for .NET (C#)
// AI Prompts: Configure the ShowValuesSetting of a pivot table so that all data fields use a largest‑to‑smallest ranking calculation in Aspose.Cells (C#). | Apply a uniform ranking calculation to the Amount and Quantity data fields of a pivot table using the Aspose.Cells API.
// Common Searches: Aspose.Cells C# set ShowValuesSetting calculation type for multiple pivot data fields | Rank largest to smallest for pivot table values using Aspose.Cells .NET | How to apply the same ShowValuesSetting to all data items in an Aspose.Cells pivot table
// Tags: aspocells pivot showvalues calculationtype ranklargesttosmallest | c# set pivot data field ranking aspocells | uniform showvalues setting multiple fields aspocells | pivot table ranking calculation aspocells c#

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // The example creates a workbook, fills it with sample data, adds a pivot table with Category rows, SubCategory columns, and Amount and Quantity as data fields, sets ShowValuesSetting.CalculationType to RankLargestToSmallest for both data fields, refreshes and calculates the pivot, and saves the result as PivotShowValues_RankLargestToSmallest.xlsx.
    public class SetShowValuesCalculationForMultipleFields
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the pivot table
                sheet.Cells["A1"].Value = "Category";
                sheet.Cells["B1"].Value = "SubCategory";
                sheet.Cells["C1"].Value = "Amount";
                sheet.Cells["D1"].Value = "Quantity";

                string[] categories = { "A", "A", "B", "B", "C", "C" };
                string[] subCategories = { "X", "Y", "X", "Y", "X", "Y" };
                double[] amounts = { 1000, 1500, 2000, 2500, 3000, 3500 };
                int[] quantities = { 10, 15, 20, 25, 30, 35 };

                for (int i = 0; i < categories.Length; i++)
                {
                    int row = i + 2; // Data starts from row 2
                    sheet.Cells[row, 0].Value = categories[i];
                    sheet.Cells[row, 1].Value = subCategories[i];
                    sheet.Cells[row, 2].Value = amounts[i];
                    sheet.Cells[row, 3].Value = quantities[i];
                }

                // Add a pivot table based on the data range
                int pivotIndex = sheet.PivotTables.Add("A1:D7", "F3", "PivotTable1");
                PivotTable pivotTable = sheet.PivotTables[pivotIndex];

                // Add fields to the pivot table
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
                pivotTable.AddFieldToArea(PivotFieldType.Column, "SubCategory");

                // Add two data fields: Amount and Quantity
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Quantity");

                // Refresh and calculate the pivot table data using the correct API
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Save the workbook
                workbook.Save("PivotShowValues_RankLargestToSmallest.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            SetShowValuesCalculationForMultipleFields.Run();
        }
    }
}
