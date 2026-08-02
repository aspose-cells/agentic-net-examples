using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotCloneDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook and add sample data
            Workbook workbook = new Workbook();
            Worksheet dataSheet = workbook.Worksheets[0];
            dataSheet.Name = "Data";

            // Populate sample data for the pivot table
            dataSheet.Cells["A1"].PutValue("Category");
            dataSheet.Cells["B1"].PutValue("Product");
            dataSheet.Cells["C1"].PutValue("Sales");

            dataSheet.Cells["A2"].PutValue("Fruit");
            dataSheet.Cells["B2"].PutValue("Apple");
            dataSheet.Cells["C2"].PutValue(120);

            dataSheet.Cells["A3"].PutValue("Fruit");
            dataSheet.Cells["B3"].PutValue("Banana");
            dataSheet.Cells["C3"].PutValue(80);

            dataSheet.Cells["A4"].PutValue("Vegetable");
            dataSheet.Cells["B4"].PutValue("Carrot");
            dataSheet.Cells["C4"].PutValue(150);

            // Create the source pivot table on the same sheet
            PivotTableCollection sourcePivots = dataSheet.PivotTables;
            int sourcePivotIndex = sourcePivots.Add("=Data!A1:C4", "E1", "SourcePivot");
            PivotTable sourcePivot = sourcePivots[sourcePivotIndex];

            // Configure the source pivot table (layout & formatting)
            sourcePivot.AddFieldToArea(PivotFieldType.Row, "Category");
            sourcePivot.AddFieldToArea(PivotFieldType.Column, "Product");
            sourcePivot.AddFieldToArea(PivotFieldType.Data, "Sales");
            sourcePivot.PivotTableStyleType = PivotTableStyleType.PivotTableStyleMedium9;
            sourcePivot.ShowInTabularForm();

            // Refresh to calculate data
            sourcePivot.CalculateData();

            // Add a new worksheet where the cloned pivot table will reside
            Worksheet cloneSheet = workbook.Worksheets.Add("ClonedPivotSheet");

            // Clone the existing pivot table to the new worksheet
            // This method copies layout, formatting, and style from the source pivot table
            int clonedPivotIndex = cloneSheet.PivotTables.Add(sourcePivot, "A1", "ClonedPivot");
            PivotTable clonedPivot = cloneSheet.PivotTables[clonedPivotIndex];

            // Refresh the cloned pivot table to populate its data
            clonedPivot.CalculateData();

            // Optionally refresh all pivot tables in the workbook (ensures consistency)
            foreach (Worksheet ws in workbook.Worksheets)
            {
                ws.RefreshPivotTables();
            }

            // Save the workbook with the cloned pivot table
            workbook.Save("ClonedPivotTableDemo.xlsx");
        }
    }
}