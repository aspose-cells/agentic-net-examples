using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotUnionExample
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // -------------------------------------------------
                // Prepare three worksheets with sample data
                // -------------------------------------------------
                // Sheet1
                Worksheet sheet1 = workbook.Worksheets[0];
                sheet1.Name = "Sheet1";
                FillSampleData(sheet1);

                // Sheet2
                Worksheet sheet2 = workbook.Worksheets.Add("Sheet2");
                FillSampleData(sheet2);

                // Sheet3
                Worksheet sheet3 = workbook.Worksheets.Add("Sheet3");
                FillSampleData(sheet3);

                // -------------------------------------------------
                // Add a worksheet that will contain the PivotTable
                // -------------------------------------------------
                Worksheet pivotSheet = workbook.Worksheets.Add("PivotSheet");
                PivotTableCollection pivotTables = pivotSheet.PivotTables;

                // Define the consolidation (union) ranges from the three sheets
                string[] sourceRanges = new string[]
                {
                    "Sheet1!A1:C5",
                    "Sheet2!A1:C5",
                    "Sheet3!A1:C5"
                };

                // Add the PivotTable using the consolidation ranges overload.
                // Pass null for page fields because they are not used in this example.
                int pivotIndex = pivotTables.Add(sourceRanges, false, null, "A1", "ConsolidatedPivot");
                PivotTable pivotTable = pivotTables[pivotIndex];

                // Configure the PivotTable (Category as Row, Year as Column, Amount as Data)
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
                pivotTable.AddFieldToArea(PivotFieldType.Column, "Year");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

                // Refresh and calculate the PivotTable data
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Save the workbook
                workbook.Save("ConsolidatedPivotTable.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Helper method to fill each worksheet with identical sample data
        private static void FillSampleData(Worksheet sheet)
        {
            Cells cells = sheet.Cells;
            // Headers
            cells["A1"].PutValue("Category");
            cells["B1"].PutValue("Year");
            cells["C1"].PutValue("Amount");

            // Sample rows
            string[] categories = { "Food", "Food", "Beverage", "Beverage", "Misc" };
            int[] years = { 2020, 2021, 2020, 2021, 2020 };
            int[] amounts = { 120, 150, 80, 95, 60 };

            for (int i = 0; i < categories.Length; i++)
            {
                int row = i + 2; // Data starts at row 2
                cells[$"A{row}"].PutValue(categories[i]);
                cells[$"B{row}"].PutValue(years[i]);
                cells[$"C{row}"].PutValue(amounts[i]);
            }
        }
    }
}