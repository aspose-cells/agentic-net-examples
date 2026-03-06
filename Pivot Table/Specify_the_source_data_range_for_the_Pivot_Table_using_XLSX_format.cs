using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsPivotSourceExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // -------------------------------------------------
            // 1. Prepare source data on the first worksheet
            // -------------------------------------------------
            Worksheet dataSheet = workbook.Worksheets[0];
            dataSheet.Name = "Data";

            // Fill sample data (A1:B5)
            dataSheet.Cells["A1"].PutValue("Category");
            dataSheet.Cells["B1"].PutValue("Amount");
            dataSheet.Cells["A2"].PutValue("Food");
            dataSheet.Cells["B2"].PutValue(1200);
            dataSheet.Cells["A3"].PutValue("Clothing");
            dataSheet.Cells["B3"].PutValue(800);
            dataSheet.Cells["A4"].PutValue("Electronics");
            dataSheet.Cells["B4"].PutValue(1500);
            dataSheet.Cells["A5"].PutValue("Travel");
            dataSheet.Cells["B5"].PutValue(700);

            // -------------------------------------------------
            // 2. Create a worksheet that will host the PivotTable
            // -------------------------------------------------
            Worksheet pivotSheet = workbook.Worksheets.Add("Pivot");

            // -------------------------------------------------
            // 3. Define the source data range in XLSX format
            // -------------------------------------------------
            AsposeRange usedRange = dataSheet.Cells.MaxDisplayRange;

            int startRow = usedRange.FirstRow;
            int startColumn = usedRange.FirstColumn;
            int endRow = startRow + usedRange.RowCount - 1;
            int endColumn = startColumn + usedRange.ColumnCount - 1;

            string startAddress = CellsHelper.CellIndexToName(startRow, startColumn);
            string endAddress = CellsHelper.CellIndexToName(endRow, endColumn);
            string sourceData = $"={dataSheet.Name}!{startAddress}:{endAddress}";

            // -------------------------------------------------
            // 4. Add the PivotTable using the Add(string, string, string) overload
            // -------------------------------------------------
            string destCell = "A1";
            string pivotName = "SalesPivot";

            int pivotIndex = pivotSheet.PivotTables.Add(sourceData, destCell, pivotName);
            PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

            // -------------------------------------------------
            // 5. Configure the PivotTable (optional)
            // -------------------------------------------------
            // Add "Category" as Row field (field index 0)
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);
            // Add "Amount" as Data field (field index 1)
            pivotTable.AddFieldToArea(PivotFieldType.Data, 1);

            // Refresh data and calculate the report
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // -------------------------------------------------
            // 6. Verify the source data range
            // -------------------------------------------------
            string[] sourceInfo = pivotTable.GetSource();
            Console.WriteLine("Pivot Table Source Range: " + sourceInfo[0]);

            // -------------------------------------------------
            // 7. Save the workbook in XLSX format
            // -------------------------------------------------
            workbook.Save("PivotTableWithSource.xlsx");
        }
    }
}