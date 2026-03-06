using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotSortExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            worksheet.Cells["A1"].PutValue("Country");
            worksheet.Cells["B1"].PutValue("Sales");
            worksheet.Cells["A2"].PutValue("USA");
            worksheet.Cells["A3"].PutValue("UK");
            worksheet.Cells["A4"].PutValue("Germany");
            worksheet.Cells["B2"].PutValue(1000);
            worksheet.Cells["B3"].PutValue(2000);
            worksheet.Cells["B4"].PutValue(1500);

            // Add a pivot table based on the data range A1:B4, place it at E3
            int pivotIndex = worksheet.PivotTables.Add("A1:B4", "E3", "SalesPivot");
            PivotTable pivotTable = worksheet.PivotTables[pivotIndex];

            // Add the "Country" field to the row area
            int rowFieldIndex = pivotTable.AddFieldToArea(PivotFieldType.Row, "Country");
            PivotField rowField = pivotTable.RowFields[rowFieldIndex];

            // Add the "Sales" field to the data area
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

            // Sort the row field in descending order by its own data labels (-1)
            rowField.SortBy(SortOrder.Descending, -1);

            // Refresh the pivot table data and calculate the results
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook in XLSX format
            workbook.Save("SortedPivotTable.xlsx");
        }
    }
}