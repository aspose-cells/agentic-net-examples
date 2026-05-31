using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook instance
        Workbook workbook = new Workbook();

        // Add a worksheet that will be used for the pivot table
        Worksheet pivotWorksheet = workbook.Worksheets.Add("PivotData");

        // (Optional) Populate the worksheet with sample data
        pivotWorksheet.Cells["A1"].PutValue("Category");
        pivotWorksheet.Cells["B1"].PutValue("Amount");
        pivotWorksheet.Cells["A2"].PutValue("Food");
        pivotWorksheet.Cells["B2"].PutValue(120);
        pivotWorksheet.Cells["A3"].PutValue("Travel");
        pivotWorksheet.Cells["B3"].PutValue(80);

        // Save the workbook (optional)
        workbook.Save("PivotWorkbook.xlsx", SaveFormat.Xlsx);
    }
}