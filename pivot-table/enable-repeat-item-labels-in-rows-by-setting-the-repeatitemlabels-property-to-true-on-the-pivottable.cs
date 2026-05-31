using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

class EnableRepeatItemLabelsDemo
{
    public static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add sample data for the pivot table
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["A4"].PutValue("C");
        sheet.Cells["B4"].PutValue(30);

        // Create a pivot table
        int pivotIndex = sheet.PivotTables.Add("A1:B4", "D1", "PivotTable1");
        PivotTable pivotTable = sheet.PivotTables[pivotIndex];

        // Add a row field and a data field
        pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
        pivotTable.AddFieldToArea(PivotFieldType.Data, "Value");

        // Enable repeat item labels for each row field
        foreach (PivotField rowField in pivotTable.RowFields)
        {
            rowField.IsRepeatItemLabels = true;
        }

        // Populate the pivot table with calculated data
        pivotTable.CalculateData();

        // Save the workbook
        workbook.Save("RepeatItemLabelsDemo.xlsx");
    }
}