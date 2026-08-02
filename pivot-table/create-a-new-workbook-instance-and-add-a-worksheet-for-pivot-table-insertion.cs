using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook instance
        Workbook workbook = new Workbook();

        // Add a worksheet that will hold the data for the pivot table
        Worksheet pivotWorksheet = workbook.Worksheets.Add("PivotData");

        // (Optional) Populate the worksheet with sample data
        // pivotWorksheet.Cells["A1"].PutValue("Category");
        // pivotWorksheet.Cells["B1"].PutValue("Amount");
        // pivotWorksheet.Cells["A2"].PutValue("A");
        // pivotWorksheet.Cells["B2"].PutValue(100);
        // pivotWorksheet.Cells["A3"].PutValue("B");
        // pivotWorksheet.Cells["B3"].PutValue(200);

        // Save the workbook to a file
        workbook.Save("PivotWorkbook.xlsx", SaveFormat.Xlsx);
    }
}