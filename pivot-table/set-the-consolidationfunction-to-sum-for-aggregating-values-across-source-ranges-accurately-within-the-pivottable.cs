using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    class SetPivotFieldFunctionSum
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate sample data for the pivot table
                cells["A1"].Value = "Product";
                cells["B1"].Value = "Sales";
                cells["A2"].Value = "A";
                cells["B2"].Value = 100;
                cells["A3"].Value = "B";
                cells["B3"].Value = 120;
                cells["A4"].Value = "A";
                cells["B4"].Value = 80;
                cells["A5"].Value = "B";
                cells["B5"].Value = 60;

                // Add a pivot table to the worksheet
                PivotTableCollection pivotTables = sheet.PivotTables;
                int pivotIndex = pivotTables.Add("A1:B5", "E3", "PivotTable1");
                PivotTable pivotTable = pivotTables[pivotIndex];

                // Add a row field (Product)
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");

                // Add a data field (Sales) and set its consolidation function to Sum
                int dataFieldIdx = pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");
                PivotField dataField = pivotTable.DataFields[dataFieldIdx];
                dataField.Function = ConsolidationFunction.Sum;

                // Refresh the pivot table data and calculate the results
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Save the workbook with the configured pivot table
                workbook.Save("PivotFieldFunctionSum.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            SetPivotFieldFunctionSum.Run();
        }
    }
}