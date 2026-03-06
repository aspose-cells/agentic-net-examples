using System;
using Aspose.Cells;
using Aspose.Cells.Tables;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    public class ConsolidationFunctionsDemo
    {
        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // -------------------------------------------------
            // 1. Populate sample data (Category | Value)
            // -------------------------------------------------
            cells["A1"].PutValue("Category");
            cells["B1"].PutValue("Value");
            string[] categories = { "A", "B", "A", "B", "A", "B" };
            int[] values = { 10, 20, 30, 40, 50, 60 };
            for (int i = 0; i < categories.Length; i++)
            {
                cells[i + 1, 0].PutValue(categories[i]);   // Column A
                cells[i + 1, 1].PutValue(values[i]);      // Column B
            }

            // -------------------------------------------------
            // 2. Apply Subtotals (Sum, Count, Average, Max, Min)
            // -------------------------------------------------
            // Define the data range for subtotals (A1:B7)
            CellArea dataArea = CellArea.CreateCellArea(0, 0, categories.Length, 1);

            // Group by the first column (Category) and apply each function to the second column (Value)
            cells.Subtotal(dataArea, 0, ConsolidationFunction.Sum, new int[] { 1 });
            cells.Subtotal(dataArea, 0, ConsolidationFunction.Count, new int[] { 1 });
            cells.Subtotal(dataArea, 0, ConsolidationFunction.Average, new int[] { 1 });
            cells.Subtotal(dataArea, 0, ConsolidationFunction.Max, new int[] { 1 });
            cells.Subtotal(dataArea, 0, ConsolidationFunction.Min, new int[] { 1 });

            // -------------------------------------------------
            // 3. Create a Table and set Totals row calculations
            // -------------------------------------------------
            // Add a ListObject (table) covering the same range
            ListObjectCollection tables = sheet.ListObjects;
            int tableIndex = tables.Add(0, 0, categories.Length + 1, 1, true);
            ListObject table = tables[tableIndex];
            table.ShowTotals = true; // Enable totals row

            // Set TotalsCalculation for the "Value" column (index 1)
            ListColumn valueColumn = table.ListColumns[1];
            valueColumn.TotalsCalculation = TotalsCalculation.Sum;
            valueColumn.TotalsCalculation = TotalsCalculation.Count;
            valueColumn.TotalsCalculation = TotalsCalculation.Average;
            valueColumn.TotalsCalculation = TotalsCalculation.Max;
            valueColumn.TotalsCalculation = TotalsCalculation.Min;

            // -------------------------------------------------
            // 4. Create PivotTables and set the Function property
            // -------------------------------------------------
            void CreatePivotWithFunction(ConsolidationFunction func, string pivotName, int startRow)
            {
                int pivotIdx = sheet.PivotTables.Add("A1:B7", $"E{startRow}", pivotName);
                PivotTable pivot = sheet.PivotTables[pivotIdx];

                pivot.AddFieldToArea(PivotFieldType.Row, "Category");
                int dataFieldIdx = pivot.AddFieldToArea(PivotFieldType.Data, "Value");
                PivotField dataField = pivot.DataFields[dataFieldIdx];
                dataField.Function = func;

                pivot.RefreshData();
                pivot.CalculateData();
            }

            CreatePivotWithFunction(ConsolidationFunction.Sum, "Pivot_Sum", 12);
            CreatePivotWithFunction(ConsolidationFunction.Count, "Pivot_Count", 22);
            CreatePivotWithFunction(ConsolidationFunction.Average, "Pivot_Average", 32);
            CreatePivotWithFunction(ConsolidationFunction.Max, "Pivot_Max", 42);
            CreatePivotWithFunction(ConsolidationFunction.Min, "Pivot_Min", 52);

            // -------------------------------------------------
            // 5. Save the workbook to XLSX format
            // -------------------------------------------------
            workbook.Save("ConsolidationFunctionsDemo.xlsx");
        }
    }

    public class Program
    {
        public static void Main()
        {
            ConsolidationFunctionsDemo.Run();
        }
    }
}