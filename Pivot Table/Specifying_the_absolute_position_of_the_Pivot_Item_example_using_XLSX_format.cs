using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    public class PivotItemAbsolutePositionDemo
    {
        public static void Run()
        {
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["A4"].PutValue("Orange");
            sheet.Cells["B2"].PutValue(1000);
            sheet.Cells["B3"].PutValue(2000);
            sheet.Cells["B4"].PutValue(3000);

            int pivotIndex = sheet.PivotTables.Add("A1:B4", "E3", "ProductPivot");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

            pivotTable.RefreshData();
            pivotTable.CalculateData();

            PivotField rowField = pivotTable.RowFields[0];
            PivotItemCollection items = rowField.PivotItems;

            items["Apple"].Position = 2;
            items["Banana"].Position = 0;
            items["Orange"].Position = 1;

            pivotTable.CalculateData();

            workbook.Save("PivotItemAbsolutePositionDemo.xlsx");
        }

        public static void Main(string[] args)
        {
            Run();
        }
    }
}