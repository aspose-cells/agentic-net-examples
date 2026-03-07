using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

class RenderSlicerExample
{
    static void Main()
    {
        // Paths for input and output files
        string inputPath = "sample.xlsx";
        string outputPath = "sample_with_slicer.xlsx";

        // Load the existing workbook (lifecycle: load)
        Workbook workbook = new Workbook(inputPath);

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Obtain a pivot table; if none exists, create a simple one
        PivotTable pivot;
        if (sheet.PivotTables.Count == 0)
        {
            // Create sample data for the pivot table
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Amount");
            sheet.Cells["A2"].PutValue("Fruit");
            sheet.Cells["B2"].PutValue(100);
            sheet.Cells["A3"].PutValue("Vegetable");
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["A4"].PutValue("Fruit");
            sheet.Cells["B4"].PutValue(200);
            sheet.Cells["A5"].PutValue("Vegetable");
            sheet.Cells["B5"].PutValue(120);

            // Add a pivot table based on the sample data
            int pivotIndex = sheet.PivotTables.Add("A1:B5", "D1", "PivotTable1");
            pivot = sheet.PivotTables[pivotIndex];
            pivot.AddFieldToArea(PivotFieldType.Row, "Category");
            pivot.AddFieldToArea(PivotFieldType.Data, "Amount");
        }
        else
        {
            // Use the first existing pivot table
            pivot = sheet.PivotTables[0];
        }

        // Add a slicer linked to the first base field of the pivot table
        // Row and column specify the upper‑left corner of the slicer on the sheet
        int slicerIndex = sheet.Slicers.Add(pivot, 2, 5, 0);
        Slicer slicer = sheet.Slicers[slicerIndex];

        // Optional: give the slicer a meaningful name
        slicer.Name = "CategorySlicer";

        // Save the modified workbook (lifecycle: save)
        workbook.Save(outputPath, SaveFormat.Xlsx);
    }
}