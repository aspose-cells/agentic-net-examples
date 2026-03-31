using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

namespace AsposeCellsSlicerDemo
{
    public class Program
    {
        public static void Main()
        {
            // Load an existing XLSX workbook
            Workbook workbook = new Workbook("input.xlsx");

            // Get the first worksheet (or any target worksheet)
            Worksheet sheet = workbook.Worksheets[0];

            // -------------------------------------------------
            // 1. Create a PivotTable from existing data range
            // -------------------------------------------------
            // Assume the source data is in A1:B4 of the same sheet
            // Destination cell for the PivotTable report is C3
            // PivotTable name is "MyPivot"
            int pivotIndex = sheet.PivotTables.Add("A1:B4", "C3", "MyPivot");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Configure the PivotTable (e.g., Row field = first column, Data field = second column)
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Column A as Row field
            pivotTable.AddFieldToArea(PivotFieldType.Data, 1); // Column B as Data field

            // -------------------------------------------------
            // 2. Add a Slicer linked to the PivotTable
            // -------------------------------------------------
            // Add slicer at top‑left corner cell (row 0, column 0) using the Row field name "Fruit"
            // Adjust the field name to match the actual header in column A
            int slicerIndex = sheet.Slicers.Add(pivotTable, 0, 0, "Fruit");
            Slicer slicer = sheet.Slicers[slicerIndex];

            // -------------------------------------------------
            // 3. Connect the slicer to the PivotTable
            // -------------------------------------------------
            slicer.AddPivotConnection(pivotTable);

            // Optional: refresh the slicer to ensure it reflects the current PivotTable state
            slicer.Refresh();

            // Save the modified workbook
            workbook.Save("output.xlsx", SaveFormat.Xlsx);
        }
    }
}