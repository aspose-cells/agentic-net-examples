using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotPaginationDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            sheet.Cells["A1"].Value = "Category";
            sheet.Cells["B1"].Value = "Region";
            sheet.Cells["C1"].Value = "Sales";

            sheet.Cells["A2"].Value = "Food";
            sheet.Cells["B2"].Value = "North";
            sheet.Cells["C2"].Value = 1200;

            sheet.Cells["A3"].Value = "Food";
            sheet.Cells["B3"].Value = "South";
            sheet.Cells["C3"].Value = 1500;

            sheet.Cells["A4"].Value = "Drink";
            sheet.Cells["B4"].Value = "North";
            sheet.Cells["C4"].Value = 800;

            sheet.Cells["A5"].Value = "Drink";
            sheet.Cells["B5"].Value = "South";
            sheet.Cells["C5"].Value = 950;

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:C5", "E3", "PivotTable1");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Add fields to the pivot table
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
            pivotTable.AddFieldToArea(PivotFieldType.Column, "Region");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

            // Add a page (report filter) field – this will be used for pagination
            pivotTable.AddFieldToArea(PivotFieldType.Page, "Category");

            // Refresh and calculate the pivot table to ensure data is up‑to‑date
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Enable pagination by generating a separate printed page for each page field item
            // ShowReportFilterPage creates a separate worksheet page for each item of the page field
            foreach (PivotField pageField in pivotTable.PageFields)
            {
                pivotTable.ShowReportFilterPage(pageField);
            }

            // Save the workbook as ODS (OpenDocument Spreadsheet) format
            workbook.Save("PivotTable_Pagination_Output.ods", SaveFormat.Ods);
        }
    }
}