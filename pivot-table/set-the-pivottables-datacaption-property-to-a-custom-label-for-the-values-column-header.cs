using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    public class SetPivotTableDataCaptionDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the pivot table
                sheet.Cells["A1"].Value = "Product";
                sheet.Cells["B1"].Value = "Region";
                sheet.Cells["C1"].Value = "Sales";

                sheet.Cells["A2"].Value = "Laptop";
                sheet.Cells["B2"].Value = "North";
                sheet.Cells["C2"].Value = 1000;

                sheet.Cells["A3"].Value = "Laptop";
                sheet.Cells["B3"].Value = "South";
                sheet.Cells["C3"].Value = 1500;

                sheet.Cells["A4"].Value = "Phone";
                sheet.Cells["B4"].Value = "North";
                sheet.Cells["C4"].Value = 800;

                sheet.Cells["A5"].Value = "Phone";
                sheet.Cells["B5"].Value = "South";
                sheet.Cells["C5"].Value = 1200;

                // Add a pivot table to the worksheet
                int pivotIndex = sheet.PivotTables.Add("A1:C5", "E3", "PivotTable1");
                PivotTable pivotTable = sheet.PivotTables[pivotIndex];

                // Configure the pivot table fields
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
                pivotTable.AddFieldToArea(PivotFieldType.Column, "Region");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

                // Set a custom caption for the values column header (DataFieldHeaderName)
                pivotTable.DataFieldHeaderName = "Custom Values";

                // Refresh and calculate the pivot table to apply changes
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Save the workbook
                workbook.Save("PivotTableWithCustomDataCaption.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            SetPivotTableDataCaptionDemo.Run();
        }
    }
}