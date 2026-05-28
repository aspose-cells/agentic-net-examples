using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    public class SetRankLargestToSmallestForMultipleDataFields
    {
        public static void Main()
        {
            Run();
        }

        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data
                sheet.Cells["A1"].Value = "Category";
                sheet.Cells["B1"].Value = "Sales";
                sheet.Cells["C1"].Value = "Quantity";

                sheet.Cells["A2"].Value = "A";
                sheet.Cells["B2"].Value = 1200;
                sheet.Cells["C2"].Value = 30;

                sheet.Cells["A3"].Value = "B";
                sheet.Cells["B3"].Value = 1500;
                sheet.Cells["C3"].Value = 45;

                sheet.Cells["A4"].Value = "C";
                sheet.Cells["B4"].Value = 800;
                sheet.Cells["C4"].Value = 20;

                sheet.Cells["A5"].Value = "A";
                sheet.Cells["B5"].Value = 1100;
                sheet.Cells["C5"].Value = 25;

                sheet.Cells["A6"].Value = "B";
                sheet.Cells["B6"].Value = 1300;
                sheet.Cells["C6"].Value = 35;

                // Add a pivot table based on the data range
                int pivotIndex = sheet.PivotTables.Add("A1:C6", "E3", "PivotTable1");
                PivotTable pivotTable = sheet.PivotTables[pivotIndex];

                // Add a row field (Category)
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");

                // Add two data fields (Sales and Quantity)
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Quantity");

                // Set ranking calculation for each data field
                foreach (PivotField dataField in pivotTable.DataFields)
                {
                    dataField.ShowValuesSetting.CalculationType = PivotFieldDataDisplayFormat.RankLargestToSmallest;
                }

                // Refresh and calculate the pivot table data
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Save the workbook
                string outputPath = "Pivot_RankLargestToSmallest.xlsx";
                workbook.Save(outputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}