using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

namespace AsposeCellsSlicerVersioningDemo
{
    public class Program
    {
        public static void Main()
        {
            // Initialize a new workbook
            Workbook workbook = new Workbook();

            // -------------------------------------------------
            // 1. Prepare sample data for the pivot table
            // -------------------------------------------------
            Worksheet dataSheet = workbook.Worksheets[0];
            dataSheet.Name = "Data";

            // Header row
            dataSheet.Cells["A1"].PutValue("Category");
            dataSheet.Cells["B1"].PutValue("Amount");

            // Sample rows
            dataSheet.Cells["A2"].PutValue("Food");
            dataSheet.Cells["B2"].PutValue(120);
            dataSheet.Cells["A3"].PutValue("Beverage");
            dataSheet.Cells["B3"].PutValue(80);
            dataSheet.Cells["A4"].PutValue("Food");
            dataSheet.Cells["B4"].PutValue(150);
            dataSheet.Cells["A5"].PutValue("Beverage");
            dataSheet.Cells["B5"].PutValue(70);
            dataSheet.Cells["A6"].PutValue("Stationery");
            dataSheet.Cells["B6"].PutValue(40);

            // -------------------------------------------------
            // 2. Create a pivot table based on the data
            // -------------------------------------------------
            Worksheet pivotSheet = workbook.Worksheets.Add("Pivot");
            int pivotIndex = pivotSheet.PivotTables.Add("A1:B6", "C3", "SalesPivot");
            PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

            // Configure pivot fields
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Category
            pivotTable.AddFieldToArea(PivotFieldType.Data, 1); // Sum of Amount

            // -------------------------------------------------
            // 3. Add a slicer linked to the pivot table
            // -------------------------------------------------
            Worksheet slicerSheet = workbook.Worksheets.Add("Slicer");
            int slicerIndex = slicerSheet.Slicers.Add(pivotTable, "A1", "Category");
            Slicer slicer = slicerSheet.Slicers[slicerIndex];

            // Initial save (version 1)
            workbook.Save("Workbook_V1.xlsx");

            // -------------------------------------------------
            // 4. First modification: change caption
            // -------------------------------------------------
            slicer.Caption = "Product Category";
            slicer.Refresh(); // Refresh slicer and underlying pivot table
            workbook.Save("Workbook_V2.xlsx");

            // -------------------------------------------------
            // 5. Second modification: lock the slicer position
            // -------------------------------------------------
            slicer.LockedPosition = true;
            slicer.Refresh();
            workbook.Save("Workbook_V3.xlsx");

            // -------------------------------------------------
            // 6. Third modification: change visual layout
            // -------------------------------------------------
            slicer.NumberOfColumns = 2;
            slicer.WidthPixel = 250;
            slicer.HeightPixel = 180;
            slicer.Refresh();
            workbook.Save("Workbook_V4.xlsx");

            // -------------------------------------------------
            // 7. Cleanup
            // -------------------------------------------------
            workbook.Dispose();

            Console.WriteLine("Workbook versions saved successfully.");
        }
    }
}