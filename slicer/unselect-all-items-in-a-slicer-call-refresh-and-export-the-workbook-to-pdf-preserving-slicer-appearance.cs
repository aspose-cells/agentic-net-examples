using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;
using Aspose.Cells.Rendering;

namespace AsposeCellsSlicerPdfExport
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data for a pivot table
            cells["A1"].Value = "Fruit";
            cells["B1"].Value = "Sales";
            cells["A2"].Value = "Apple";
            cells["B2"].Value = 120;
            cells["A3"].Value = "Banana";
            cells["B3"].Value = 80;
            cells["A4"].Value = "Orange";
            cells["B4"].Value = 150;

            // Add a pivot table based on the data
            int pivotIdx = sheet.PivotTables.Add("A1:B4", "D1", "FruitPivot");
            PivotTable pivot = sheet.PivotTables[pivotIdx];
            pivot.AddFieldToArea(PivotFieldType.Row, "Fruit");
            pivot.AddFieldToArea(PivotFieldType.Data, "Sales");
            pivot.RefreshData();
            pivot.CalculateData();

            // Add a slicer linked to the pivot table (field "Fruit")
            int slicerIdx = sheet.Slicers.Add(pivot, "F1", "Fruit");
            Slicer slicer = sheet.Slicers[slicerIdx];

            // Unselect all items: show all items (no filter) and refresh the slicer
            slicer.ShowAllItems = true;   // clears any selection
            slicer.Refresh();             // refreshes slicer and underlying pivot table

            // Prepare PDF save options – keep default sheet set (all visible sheets)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Export the workbook to PDF while preserving slicer appearance
            workbook.Save("WorkbookWithSlicer.pdf", pdfOptions);
        }
    }
}