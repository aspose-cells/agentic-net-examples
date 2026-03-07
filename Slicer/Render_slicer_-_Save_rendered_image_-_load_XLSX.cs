using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

class Program
{
    static void Main()
    {
        // *** Create a new workbook (creation rule) ***
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for a pivot table
        sheet.Cells["A1"].Value = "Category";
        sheet.Cells["B1"].Value = "Amount";
        sheet.Cells["A2"].Value = "Food";
        sheet.Cells["B2"].Value = 120;
        sheet.Cells["A3"].Value = "Transport";
        sheet.Cells["B3"].Value = 80;
        sheet.Cells["A4"].Value = "Utilities";
        sheet.Cells["B4"].Value = 150;

        // Add a pivot table based on the data range
        int pivotIdx = sheet.PivotTables.Add("A1:B4", "D3", "Pivot1");
        PivotTable pivot = sheet.PivotTables[pivotIdx];
        pivot.AddFieldToArea(PivotFieldType.Row, "Category");
        pivot.AddFieldToArea(PivotFieldType.Data, "Amount");

        // Add a slicer linked to the pivot table's "Category" field
        int slicerIdx = sheet.Slicers.Add(pivot, "F2", "Category");
        Slicer slicer = sheet.Slicers[slicerIdx];
        slicer.Name = "CategorySlicer";

        // Ensure the slicer is printable (via its shape)
        slicer.Shape.IsPrintable = true;

        // *** Render the worksheet (including the slicer) to an image file ***
        ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
        {
            ImageType = Aspose.Cells.Drawing.ImageType.Png
        };
        SheetRender sheetRender = new SheetRender(sheet, imgOptions);
        sheetRender.ToImage(0, "slicer_rendered.png");

        // *** Save the workbook (save rule) ***
        workbook.Save("slicer_demo.xlsx");

        // *** Load the saved workbook (load rule) ***
        Workbook loadedWorkbook = new Workbook("slicer_demo.xlsx");
        Console.WriteLine($"Workbook loaded successfully. Worksheet count: {loadedWorkbook.Worksheets.Count}");
    }
}