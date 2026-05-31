using System;
using Aspose.Cells;
using Aspose.Cells.Slicers;
using Aspose.Cells.Pivot;

namespace AsposeCellsSlicerNonPrintableDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for a pivot table
            sheet.Cells["A1"].Value = "Category";
            sheet.Cells["A2"].Value = "A";
            sheet.Cells["A3"].Value = "B";
            sheet.Cells["A4"].Value = "C";

            sheet.Cells["B1"].Value = "Amount";
            sheet.Cells["B2"].Value = 100;
            sheet.Cells["B3"].Value = 150;
            sheet.Cells["B4"].Value = 200;

            // Add a pivot table based on the data
            PivotTableCollection pivots = sheet.PivotTables;
            int pivotIdx = pivots.Add("A1:B4", "D1", "Pivot1");
            PivotTable pivot = pivots[pivotIdx];
            pivot.AddFieldToArea(PivotFieldType.Row, "Category");
            pivot.AddFieldToArea(PivotFieldType.Data, "Amount");
            pivot.RefreshData();
            pivot.CalculateData();

            // Add a slicer linked to the pivot table
            SlicerCollection slicers = sheet.Slicers;
            int slicerIdx = slicers.Add(pivot, "F1", "Category");
            Slicer slicer = slicers[slicerIdx];

            // Mark the slicer as non‑printable using the underlying Shape object
            // (Slicer.IsPrintable is obsolete; Shape.IsPrintable is the recommended property)
            slicer.Shape.IsPrintable = false;

            // Save the workbook (lifecycle: save)
            workbook.Save("SlicerNonPrintableDemo.xlsx");
        }
    }
}