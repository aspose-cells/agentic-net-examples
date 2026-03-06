using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    public class PivotFieldDataDisplayFormatExample
    {
        public static void Main(string[] args)
        {
            Run();
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data for the pivot table
            cells["A1"].Value = "Product";
            cells["B1"].Value = "Quarter";
            cells["C1"].Value = "Sales";

            cells["A2"].Value = "Laptop";
            cells["B2"].Value = "Q1";
            cells["C2"].Value = 1000;

            cells["A3"].Value = "Laptop";
            cells["B3"].Value = "Q2";
            cells["C3"].Value = 1200;

            cells["A4"].Value = "Phone";
            cells["B4"].Value = "Q1";
            cells["C4"].Value = 1500;

            cells["A5"].Value = "Phone";
            cells["B5"].Value = "Q2";
            cells["C5"].Value = 1800;

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:C5", "E3", "SalesPivot");
            PivotTable pivot = sheet.PivotTables[pivotIndex];

            // Add fields to the pivot table
            pivot.AddFieldToArea(PivotFieldType.Row, "Product");
            pivot.AddFieldToArea(PivotFieldType.Column, "Quarter");
            int dataFieldPos = pivot.AddFieldToArea(PivotFieldType.Data, "Sales");

            // Retrieve the data field object
            PivotField dataField = pivot.DataFields[dataFieldPos];

            // Set the display format using the recommended ShowValuesSetting property (modern approach)
            dataField.ShowValuesSetting.CalculationType = PivotFieldDataDisplayFormat.PercentageOfTotal;
            dataField.DisplayName = "Sales % of Total";

            // Refresh the pivot table data and calculate the results
            pivot.RefreshData();
            pivot.CalculateData();

            // Save the workbook in XLSX format
            workbook.Save("PivotFieldDataDisplayFormatExample.xlsx", SaveFormat.Xlsx);
        }
    }
}