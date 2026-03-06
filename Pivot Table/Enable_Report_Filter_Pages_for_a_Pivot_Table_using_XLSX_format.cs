using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    class Program
    {
        static void Main()
        {
            EnableReportFilterPagesDemo.Run();
        }
    }

    public class EnableReportFilterPagesDemo
    {
        public static void Run()
        {
            // Create a new workbook and add sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Headers
            cells["A1"].PutValue("Category");
            cells["B1"].PutValue("Product");
            cells["C1"].PutValue("Amount");

            // Sample rows
            cells["A2"].PutValue("Fruit");
            cells["B2"].PutValue("Apple");
            cells["C2"].PutValue(10);

            cells["A3"].PutValue("Fruit");
            cells["B3"].PutValue("Banana");
            cells["C3"].PutValue(15);

            cells["A4"].PutValue("Vegetable");
            cells["B4"].PutValue("Carrot");
            cells["C4"].PutValue(7);

            cells["A5"].PutValue("Vegetable");
            cells["B5"].PutValue("Broccoli");
            cells["C5"].PutValue(12);

            // Create a pivot table based on the data range A1:C5
            PivotTableCollection pivots = sheet.PivotTables;
            int pivotIndex = pivots.Add("=A1:C5", "E1", "PivotTable1");
            PivotTable pivotTable = pivots[pivotIndex];

            // Add fields: Category as Page (Report Filter), Product as Row, Amount as Data
            pivotTable.AddFieldToArea(PivotFieldType.Page, 0);   // Category column (index 0)
            pivotTable.AddFieldToArea(PivotFieldType.Row, 1);   // Product column (index 1)
            pivotTable.AddFieldToArea(PivotFieldType.Data, 2);  // Amount column (index 2)

            // Enable report filter pages for each page field
            foreach (PivotField pageField in pivotTable.PageFields)
            {
                pivotTable.ShowReportFilterPage(pageField);
            }

            // Save the workbook
            workbook.Save("PivotReportFilterPages.xlsx");
        }
    }
}