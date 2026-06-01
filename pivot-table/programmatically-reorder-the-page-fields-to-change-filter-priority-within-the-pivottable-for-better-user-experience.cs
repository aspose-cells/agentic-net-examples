using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotReorderPageFields
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate sample data
                cells["A1"].Value = "Category";
                cells["B1"].Value = "Region";
                cells["C1"].Value = "Sales";

                for (int i = 2; i <= 11; i++)
                {
                    cells[$"A{i}"].Value = $"Category{(i % 3) + 1}";
                    cells[$"B{i}"].Value = $"Region{(i % 2) + 1}";
                    cells[$"C{i}"].Value = i * 100;
                }

                // Add a pivot table based on the data range
                PivotTableCollection pivots = sheet.PivotTables;
                int pivotIndex = pivots.Add("A1:C11", "E3", "SalesPivot");
                PivotTable pivot = pivots[pivotIndex];

                // Add page (report filter) fields
                pivot.AddFieldToArea(PivotFieldType.Page, "Category");
                pivot.AddFieldToArea(PivotFieldType.Page, "Region");

                // Add a row field and a data field for completeness
                pivot.AddFieldToArea(PivotFieldType.Row, "Category");
                pivot.AddFieldToArea(PivotFieldType.Data, "Sales");

                // Refresh the pivot table so that fields are materialized
                pivot.RefreshData();
                pivot.CalculateData();

                // Reorder page fields: move "Region" (currently at index 1) to index 0
                if (pivot.PageFields.Count > 1)
                {
                    pivot.PageFields.Move(1, 0);
                }

                // Optional layout settings
                pivot.PageFieldOrder = PrintOrderType.DownThenOver;
                pivot.PageFieldWrapCount = 1;

                // Save the workbook
                workbook.Save("ReorderedPageFieldsPivot.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}