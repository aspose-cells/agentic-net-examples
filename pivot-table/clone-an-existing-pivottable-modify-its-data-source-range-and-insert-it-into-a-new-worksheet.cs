using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotCloneExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sourceSheet = workbook.Worksheets[0];
            sourceSheet.Name = "SourceData";

            // Populate sample data for the source pivot table
            Cells srcCells = sourceSheet.Cells;
            srcCells["A1"].PutValue("Category");
            srcCells["B1"].PutValue("Product");
            srcCells["C1"].PutValue("Sales");

            srcCells["A2"].PutValue("Fruit");
            srcCells["B2"].PutValue("Apple");
            srcCells["C2"].PutValue(120);

            srcCells["A3"].PutValue("Fruit");
            srcCells["B3"].PutValue("Banana");
            srcCells["C3"].PutValue(80);

            srcCells["A4"].PutValue("Vegetable");
            srcCells["B4"].PutValue("Carrot");
            srcCells["C4"].PutValue(150);

            // Create a pivot table on the source sheet
            string sourceDataRange = "A1:C4";
            PivotTableCollection srcPivots = sourceSheet.PivotTables;
            int srcPivotIndex = srcPivots.Add(sourceDataRange, "E1", "SourcePivot");
            PivotTable sourcePivot = srcPivots[srcPivotIndex];
            sourcePivot.AddFieldToArea(PivotFieldType.Row, "Category");
            sourcePivot.AddFieldToArea(PivotFieldType.Row, "Product");
            sourcePivot.AddFieldToArea(PivotFieldType.Data, "Sales");
            sourcePivot.CalculateData();

            // Add a new worksheet where the cloned pivot table will be placed
            Worksheet destSheet = workbook.Worksheets.Add("ClonedData");

            // Populate a different data range on the destination sheet (new source for the cloned pivot)
            Cells destCells = destSheet.Cells;
            destCells["A1"].PutValue("Category");
            destCells["B1"].PutValue("Product");
            destCells["C1"].PutValue("Sales");

            destCells["A2"].PutValue("Fruit");
            destCells["B2"].PutValue("Apple");
            destCells["C2"].PutValue(200);

            destCells["A3"].PutValue("Fruit");
            destCells["B3"].PutValue("Banana");
            destCells["C3"].PutValue(150);

            destCells["A4"].PutValue("Vegetable");
            destCells["B4"].PutValue("Carrot");
            destCells["C4"].PutValue(180);

            // Clone the source pivot table into the new worksheet
            // The Add method clones based on an existing PivotTable
            PivotTableCollection destPivots = destSheet.PivotTables;
            int clonedPivotIndex = destPivots.Add(sourcePivot, "E1", "ClonedPivot");
            PivotTable clonedPivot = destPivots[clonedPivotIndex];

            // Change the data source of the cloned pivot table to the new range on the destination sheet
            // The source array format: {"RangeAddress", "SheetName"}
            string[] newDataSource = new string[] { "A1:C4", destSheet.Name };
            clonedPivot.ChangeDataSource(newDataSource);

            // Refresh and recalculate the cloned pivot table to reflect the new data source
            clonedPivot.RefreshData();
            clonedPivot.CalculateData();

            // Save the workbook
            workbook.Save("ClonedPivotTableExample.xlsx");
        }
    }
}