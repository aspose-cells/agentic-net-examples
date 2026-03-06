using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

namespace AsposeCellsSlicerXlsbDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Amount");
            sheet.Cells["A2"].PutValue("Food");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["A3"].PutValue("Travel");
            sheet.Cells["B3"].PutValue(80);
            sheet.Cells["A4"].PutValue("Supplies");
            sheet.Cells["B4"].PutValue(150);
            sheet.Cells["A5"].PutValue("Food");
            sheet.Cells["B5"].PutValue(90);
            sheet.Cells["A6"].PutValue("Travel");
            sheet.Cells["B6"].PutValue(60);

            // Add a pivot table based on the data range A1:B6
            int pivotIndex = sheet.PivotTables.Add("A1:B6", "D1", "PivotTable1");
            PivotTable pivot = sheet.PivotTables[pivotIndex];

            // Configure the pivot table: Category as row field, Amount as data field
            pivot.AddFieldToArea(PivotFieldType.Row, 0);   // Column 0 -> Category
            pivot.AddFieldToArea(PivotFieldType.Data, 1);  // Column 1 -> Amount

            // Add a slicer linked to the pivot table.
            // Place the slicer at row 20, column 2 (cell C20) and use the first base field (Category)
            int slicerIndex = sheet.Slicers.Add(pivot, 20, 2, 0);
            // Optionally retrieve the slicer object if further customization is needed
            // Slicer slicer = sheet.Slicers[slicerIndex];

            // Prepare XLSB save options (export all column indexes is true by default)
            XlsbSaveOptions saveOptions = new XlsbSaveOptions
            {
                ExportAllColumnIndexes = true
            };

            // Save the workbook as an XLSB file using the specified options
            workbook.Save("SlicerDemo.xlsb", saveOptions);
        }
    }
}