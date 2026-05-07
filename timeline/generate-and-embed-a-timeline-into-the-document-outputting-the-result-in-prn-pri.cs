using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsTimelinePrnDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate worksheet with sample data including a date field
            cells["A1"].Value = "Product";
            cells["B1"].Value = "Date";
            cells["C1"].Value = "Sales";

            cells["A2"].Value = "P1";
            cells["B2"].Value = new DateTime(2023, 1, 1);
            cells["C2"].Value = 120;

            cells["A3"].Value = "P2";
            cells["B3"].Value = new DateTime(2023, 1, 2);
            cells["C3"].Value = 150;

            cells["A4"].Value = "P1";
            cells["B4"].Value = new DateTime(2023, 1, 3);
            cells["C4"].Value = 200;

            // Create a pivot table
            int pivotIndex = sheet.PivotTables.Add("A1:C4", "E1", "SalesPivot");
            PivotTable pivot = sheet.PivotTables[pivotIndex];
            pivot.AddFieldToArea(PivotFieldType.Row, "Product");
            pivot.AddFieldToArea(PivotFieldType.Column, "Date");
            pivot.AddFieldToArea(PivotFieldType.Data, "Sales");
            pivot.PivotTableStyleType = PivotTableStyleType.PivotTableStyleMedium9;
            pivot.RefreshData();
            pivot.CalculateData();

            // Save the workbook as PDF
            workbook.Save("SalesTimelineOutput.pdf", SaveFormat.Pdf);
        }
    }
}