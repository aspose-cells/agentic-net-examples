using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

namespace SlicerDemo
{
    class Program
    {
        static void Main()
        {
            // Load existing workbook that contains source data
            Workbook workbook = new Workbook("Data.xlsx");
            Worksheet dataSheet = workbook.Worksheets[0]; // assume data is on the first sheet

            // Create a pivot table based on the data range (adjust the range as needed)
            // Here we assume data occupies A1:B10
            string sourceRange = "A1:B10";
            string pivotPosition = "D3";
            int pivotIndex = dataSheet.PivotTables.Add(sourceRange, pivotPosition, "MyPivot");
            PivotTable pivot = dataSheet.PivotTables[pivotIndex];

            // Add fields to the pivot table (first column as row field, second column as data field)
            pivot.AddFieldToArea(PivotFieldType.Row, 0);   // first column
            pivot.AddFieldToArea(PivotFieldType.Data, 1);  // second column
            pivot.RefreshData();
            pivot.CalculateData();

            // Add a slicer linked to the pivot table's first field (row field)
            // Destination cell for the slicer is E3
            int slicerIndex = dataSheet.Slicers.Add(pivot, "E3", "Column1"); // replace "Column1" with actual field name if known
            Slicer slicer = dataSheet.Slicers[slicerIndex];

            // Modify slicer settings
            slicer.Caption = "My Data Slicer";
            slicer.StyleType = SlicerStyleType.SlicerStyleLight2;
            slicer.NumberOfColumns = 2;
            slicer.WidthPixel = 250;
            slicer.HeightPixel = 150;
            slicer.LockedPosition = false; // allow user to move the slicer

            // Refresh the slicer to ensure it reflects the current pivot data
            slicer.Refresh();

            // Save the workbook with the slicer
            workbook.Save("Result.xlsx");
        }
    }
}