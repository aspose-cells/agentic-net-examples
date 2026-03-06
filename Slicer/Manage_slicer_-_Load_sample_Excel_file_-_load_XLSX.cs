using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

class Program
{
    static void Main()
    {
        // Load the existing Excel file (XLSX)
        string inputPath = "sample.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Obtain a pivot table; if none exists, create a simple one for demonstration
        PivotTable pivotTable;
        if (worksheet.PivotTables.Count > 0)
        {
            pivotTable = worksheet.PivotTables[0];
        }
        else
        {
            // Create sample data range A1:B5
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["B1"].PutValue("Amount");
            worksheet.Cells["A2"].PutValue("Food");
            worksheet.Cells["B2"].PutValue(120);
            worksheet.Cells["A3"].PutValue("Drink");
            worksheet.Cells["B3"].PutValue(80);
            worksheet.Cells["A4"].PutValue("Food");
            worksheet.Cells["B4"].PutValue(150);
            worksheet.Cells["A5"].PutValue("Drink");
            worksheet.Cells["B5"].PutValue(70);

            // Add a pivot table based on the sample data
            int pivotIdx = worksheet.PivotTables.Add("A1:B5", "D1", "MyPivot");
            pivotTable = worksheet.PivotTables[pivotIdx];
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Category field
            pivotTable.AddFieldToArea(PivotFieldType.Data, 1);  // Amount field
        }

        // Add a slicer linked to the "Category" field of the pivot table
        // Destination cell for the slicer is F1
        int slicerIdx = worksheet.Slicers.Add(pivotTable, "F1", "Category");
        Slicer slicer = worksheet.Slicers[slicerIdx];

        // Adjust slicer shape size (optional)
        slicer.Shape.Width = 150;
        slicer.Shape.Height = 200;

        // Save the workbook with the added slicer
        string outputPath = "sample_with_slicer.xlsx";
        workbook.Save(outputPath);
    }
}