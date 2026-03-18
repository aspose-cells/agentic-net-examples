using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    public class LocalizeSubtotalAndGrandTotal
    {
        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["B2"].PutValue(100);
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["B3"].PutValue(200);
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B4"].PutValue(150);

            // Add a pivot table based on the sample data
            int pivotIdx = sheet.PivotTables.Add("A1:B4", "D1", "MyPivot");
            PivotTable pivot = sheet.PivotTables[pivotIdx];

            // Configure pivot fields: Category as row, Value as data
            pivot.AddFieldToArea(PivotFieldType.Row, 0);   // Category
            int dataFieldIdx = pivot.AddFieldToArea(PivotFieldType.Data, 1); // Value
            // Ensure the data field uses the Sum function (default)
            pivot.DataFields[dataFieldIdx].Function = ConsolidationFunction.Sum;

            // Refresh and calculate the pivot table so that it is populated
            pivot.RefreshData();
            pivot.CalculateData();

            // Save the workbook to an XLSX file
            workbook.Save("LocalizedPivot.xlsx");
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            LocalizeSubtotalAndGrandTotal.Run();
        }
    }
}