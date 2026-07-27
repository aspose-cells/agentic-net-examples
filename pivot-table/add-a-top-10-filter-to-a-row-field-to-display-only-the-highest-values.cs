using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data for the pivot table
        worksheet.Cells["A1"].Value = "Category";
        worksheet.Cells["B1"].Value = "Sales";

        worksheet.Cells["A2"].Value = "Fruit";
        worksheet.Cells["B2"].Value = 120;

        worksheet.Cells["A3"].Value = "Vegetable";
        worksheet.Cells["B3"].Value = 80;

        worksheet.Cells["A4"].Value = "Dairy";
        worksheet.Cells["B4"].Value = 150;

        worksheet.Cells["A5"].Value = "Meat";
        worksheet.Cells["B5"].Value = 200;

        worksheet.Cells["A6"].Value = "Bakery";
        worksheet.Cells["B6"].Value = 60;

        // Create a pivot table based on the data range
        int pivotIndex = worksheet.PivotTables.Add("A1:B6", "D3", "PivotTable1");
        PivotTable pivotTable = worksheet.PivotTables[pivotIndex];

        // Add the Category field as a row field
        pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");

        // Add the Sales field as a data field
        pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

        // Apply a Top 10 filter (show top N items) on the row field.
        // Parameters:
        //   valueFieldIndex = 0   -> first data field (Sales)
        //   type = PivotFilterType.Sum -> filter based on sum of Sales
        //   isTop = true          -> show top items
        //   itemCount = 3         -> show top 3 categories
        pivotTable.BaseFields[0].FilterTop10(0, PivotFilterType.Sum, true, 3);

        // Recalculate the pivot table to apply the filter
        pivotTable.CalculateData();

        // Save the workbook
        workbook.Save("PivotTop10Filter.xlsx");
    }
}