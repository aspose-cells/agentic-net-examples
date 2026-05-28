using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    public class ExpandFirstLevelPivotDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "Data";

                // Populate sample hierarchical data
                // Category -> SubCategory -> Amount
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("SubCategory");
                sheet.Cells["C1"].PutValue("Amount");

                sheet.Cells["A2"].PutValue("Fruit");
                sheet.Cells["B2"].PutValue("Apple");
                sheet.Cells["C2"].PutValue(120);

                sheet.Cells["A3"].PutValue("Fruit");
                sheet.Cells["B3"].PutValue("Banana");
                sheet.Cells["C3"].PutValue(80);

                sheet.Cells["A4"].PutValue("Vegetable");
                sheet.Cells["B4"].PutValue("Carrot");
                sheet.Cells["C4"].PutValue(50);

                sheet.Cells["A5"].PutValue("Vegetable");
                sheet.Cells["B5"].PutValue("Potato");
                sheet.Cells["C5"].PutValue(70);

                // Add a new worksheet for the pivot table
                Worksheet pivotSheet = workbook.Worksheets.Add("Pivot");

                // Create the pivot table using the data range
                int pivotIndex = pivotSheet.PivotTables.Add("=Data!A1:C5", "A3", "PivotTable1");
                PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

                // Add hierarchical row fields (Category -> SubCategory)
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
                pivotTable.AddFieldToArea(PivotFieldType.Row, "SubCategory");

                // Add data field
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

                // Enable drilldown and show drill buttons
                pivotTable.EnableDrilldown = true;
                pivotTable.ShowDrill = true;

                // Refresh and calculate the pivot data
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Collapse all details for every row field
                foreach (PivotField rowField in pivotTable.RowFields)
                {
                    rowField.HideDetail(true);
                }

                // Expand only the first level (top‑level items) of the first row field
                PivotField firstRowField = pivotTable.RowFields[0];
                foreach (PivotItem item in firstRowField.PivotItems)
                {
                    // Ensure the top‑level items are expanded
                    item.IsDetailHidden = false;
                }

                // Save the workbook
                workbook.Save("PivotTableFirstLevelExpanded.xlsx");
                Console.WriteLine("Workbook saved successfully.");
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
            ExpandFirstLevelPivotDemo.Run();
        }
    }
}