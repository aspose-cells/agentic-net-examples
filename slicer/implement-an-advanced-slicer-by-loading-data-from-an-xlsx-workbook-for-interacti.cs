using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

namespace AdvancedSlicerDemo
{
    class Program
    {
        static void Main()
        {
            // Create workbook and add sample data
            Workbook wb = new Workbook();
            Worksheet dataSheet = wb.Worksheets[0];
            dataSheet.Name = "Data";

            dataSheet.Cells["A1"].PutValue("Product");
            dataSheet.Cells["B1"].PutValue("Year");
            dataSheet.Cells["C1"].PutValue("Amount");

            object[,] sample = {
                {"Apple",  "2020", 120},
                {"Apple",  "2021", 150},
                {"Banana", "2020", 80},
                {"Banana", "2021", 90},
                {"Cherry", "2020", 60},
                {"Cherry", "2021", 70}
            };

            for (int i = 0; i < sample.GetLength(0); i++)
            {
                dataSheet.Cells[i + 1, 0].PutValue(sample[i, 0]);
                dataSheet.Cells[i + 1, 1].PutValue(sample[i, 1]);
                dataSheet.Cells[i + 1, 2].PutValue(sample[i, 2]);
            }

            // Add a new sheet for the pivot table
            Worksheet pivotSheet = wb.Worksheets.Add("PivotSheet");

            // Define the data range for the pivot table
            string dataRange = $"'{dataSheet.Name}'!A1:C{sample.GetLength(0) + 1}";

            // Add the pivot table
            int pivotIdx = pivotSheet.PivotTables.Add(dataRange, "E5", "SalesPivot");
            PivotTable pivot = pivotSheet.PivotTables[pivotIdx];

            // Configure pivot fields
            pivot.AddFieldToArea(PivotFieldType.Row, "Product");
            pivot.AddFieldToArea(PivotFieldType.Column, "Year");
            pivot.AddFieldToArea(PivotFieldType.Data, "Amount");

            // Apply style and calculate data
            pivot.PivotTableStyleType = PivotTableStyleType.PivotTableStyleMedium9;
            pivot.RefreshData();
            pivot.CalculateData();

            // Add a slicer for the "Product" field
            int slicerIdx = pivotSheet.Slicers.Add(pivot, "A1", "Product");
            Slicer slicer = pivotSheet.Slicers[slicerIdx];
            slicer.Caption = "Product Filter";
            slicer.StyleType = SlicerStyleType.SlicerStyleLight2;
            slicer.NumberOfColumns = 1;
            slicer.Shape.WidthPt = 150;
            slicer.Shape.HeightPt = 200;
            slicer.Refresh();

            // Save the workbook
            wb.Save("AdvancedSlicerDemo.xlsx");
        }
    }
}