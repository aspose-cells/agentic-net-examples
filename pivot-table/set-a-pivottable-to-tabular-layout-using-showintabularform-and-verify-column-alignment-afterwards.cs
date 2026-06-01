using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotTabularDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // -------------------------------------------------
            // Prepare source data for the pivot table
            // -------------------------------------------------
            Worksheet dataSheet = workbook.Worksheets[0];
            dataSheet.Name = "Data";

            // Header row
            dataSheet.Cells["A1"].PutValue("Category");
            dataSheet.Cells["B1"].PutValue("SubCategory");
            dataSheet.Cells["C1"].PutValue("Amount");

            // Sample rows
            dataSheet.Cells["A2"].PutValue("Fruit");
            dataSheet.Cells["B2"].PutValue("Apple");
            dataSheet.Cells["C2"].PutValue(120);

            dataSheet.Cells["A3"].PutValue("Fruit");
            dataSheet.Cells["B3"].PutValue("Banana");
            dataSheet.Cells["C3"].PutValue(80);

            dataSheet.Cells["A4"].PutValue("Vegetable");
            dataSheet.Cells["B4"].PutValue("Carrot");
            dataSheet.Cells["C4"].PutValue(50);

            dataSheet.Cells["A5"].PutValue("Vegetable");
            dataSheet.Cells["B5"].PutValue("Potato");
            dataSheet.Cells["C5"].PutValue(70);

            // -------------------------------------------------
            // Create a worksheet to host the pivot table
            // -------------------------------------------------
            Worksheet pivotSheet = workbook.Worksheets.Add("PivotTable");

            // Add the pivot table (source range, destination cell, name)
            int pivotIndex = pivotSheet.PivotTables.Add("=Data!A1:C5", "A1", "PivotTable1");
            PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

            // Add fields: Category and SubCategory as rows, Amount as data
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
            pivotTable.AddFieldToArea(PivotFieldType.Row, "SubCategory");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

            // -------------------------------------------------
            // Set the layout to Tabular form
            // -------------------------------------------------
            pivotTable.ShowInTabularForm();

            // Refresh and calculate to populate the view
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // -------------------------------------------------
            // Verify column alignment after applying Tabular layout
            // -------------------------------------------------
            // In Tabular layout the column headers and data columns should start at the same column index.
            // We compare the start column of the column header range with the start column of the data body range.
            CellArea columnHeaderRange = pivotTable.ColumnRange;   // Header area (if any)
            CellArea dataBodyRange = pivotTable.DataBodyRange;    // Data area

            // If there are no column fields, ColumnRange may be empty (StartColumn = -1). In that case,
            // we verify that the first data column aligns with the first row field column.
            int expectedStartColumn = dataBodyRange.StartColumn;

            if (columnHeaderRange.StartColumn != -1 && columnHeaderRange.StartColumn != expectedStartColumn)
            {
                throw new InvalidOperationException(
                    $"Column alignment mismatch: Column header starts at {columnHeaderRange.StartColumn}, " +
                    $"but data starts at {expectedStartColumn}.");
            }

            // Additional sanity check: ensure that each row in the data body has the same number of columns.
            int dataColumns = dataBodyRange.EndColumn - dataBodyRange.StartColumn + 1;
            if (dataColumns <= 0)
            {
                throw new InvalidOperationException("Data body range does not contain any columns.");
            }

            // -------------------------------------------------
            // Save the workbook
            // -------------------------------------------------
            workbook.Save("PivotTableTabularLayoutDemo.xlsx");
        }
    }
}