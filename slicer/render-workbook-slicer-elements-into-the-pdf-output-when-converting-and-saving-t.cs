using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data for the pivot table
        worksheet.Cells["A1"].PutValue("Fruit");
        worksheet.Cells["B1"].PutValue("Quantity");
        worksheet.Cells["A2"].PutValue("Apple");
        worksheet.Cells["B2"].PutValue(10);
        worksheet.Cells["A3"].PutValue("Orange");
        worksheet.Cells["B3"].PutValue(20);
        worksheet.Cells["A4"].PutValue("Banana");
        worksheet.Cells["B4"].PutValue(30);

        // Add a pivot table based on the data range
        int pivotIndex = worksheet.PivotTables.Add("A1:B4", "D1", "FruitPivot");
        PivotTable pivot = worksheet.PivotTables[pivotIndex];
        pivot.AddFieldToArea(PivotFieldType.Row, "Fruit");
        pivot.AddFieldToArea(PivotFieldType.Data, "Quantity");
        pivot.RefreshData();
        pivot.CalculateData();

        // Add a slicer linked to the "Fruit" field of the pivot table
        int slicerIndex = worksheet.Slicers.Add(pivot, "E1", "Fruit");
        Slicer slicer = worksheet.Slicers[slicerIndex];

        // Ensure the slicer is printable so it appears in the PDF output
        slicer.IsPrintable = true;

        // Save the workbook as PDF; slicer will be rendered in the PDF
        workbook.Save("SlicerOutput.pdf", SaveFormat.Pdf);
    }
}