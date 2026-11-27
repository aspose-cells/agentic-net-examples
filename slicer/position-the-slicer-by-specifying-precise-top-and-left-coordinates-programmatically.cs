// Title: Programmatically set the top and left pixel coordinates of a slicer in an Aspose.Cells .NET workbook (C#)
// Common Searches: C# Aspose.Cells how to move a slicer to a specific location on a worksheet | set slicer TopPixel and LeftPixel values using Aspose.Cells API | position pivot table slicer by pixel coordinates in Aspose.Cells for .NET | Aspose.Cells slicer placement example C# | adjust slicer location programmatically in Excel file with Aspose.Cells
// Tags: Aspose.Cells slicer positioning by pixel | C# set slicer TopPixel LeftPixel | Aspose.Cells pivot slicer location | Excel slicer placement .NET | programmatic slicer coordinates Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

// The example creates a workbook, adds sample data and a pivot table, inserts a slicer linked to the "Category" field, and then positions the slicer by assigning TopPixel = 100 and LeftPixel = 400 (with WidthPixel and HeightPixel). It also sets a caption and style before saving the file as SlicerPositionDemo.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate sample data for the pivot table
            cells["A1"].PutValue("Category");
            cells["B1"].PutValue("Product");
            cells["C1"].PutValue("Quantity");

            string[,] sampleData = new string[,]
            {
                { "Fruits", "Apple", "10" },
                { "Fruits", "Banana", "15" },
                { "Vegetables", "Carrot", "20" },
                { "Vegetables", "Tomato", "12" }
            };

            for (int i = 0; i < sampleData.GetLength(0); i++)
            {
                cells[i + 2, 0].PutValue(sampleData[i, 0]); // Category
                cells[i + 2, 1].PutValue(sampleData[i, 1]); // Product
                cells[i + 2, 2].PutValue(int.Parse(sampleData[i, 2])); // Quantity
            }

            // Add a pivot table based on the sample data
            int pivotIndex = worksheet.PivotTables.Add("A1:C5", "E1", "PivotTable1");
            PivotTable pivotTable = worksheet.PivotTables[pivotIndex];
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Quantity");
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Add a slicer linked to the "Category" field of the pivot table
            int slicerIndex = worksheet.Slicers.Add(pivotTable, "Category", "CategorySlicer");
            Slicer slicer = worksheet.Slicers[slicerIndex];

            // Position the slicer using precise pixel coordinates
            slicer.TopPixel = 100;    // Pixels from the top edge of the worksheet
            slicer.LeftPixel = 400;   // Pixels from the left edge of the worksheet
            slicer.WidthPixel = 200;  // Width of the slicer in pixels
            slicer.HeightPixel = 150; // Height of the slicer in pixels

            // Optional: set additional slicer properties
            slicer.Caption = "Category Filter";
            slicer.StyleType = SlicerStyleType.SlicerStyleLight2;

            // Save the workbook
            workbook.Save("SlicerPositionDemo.xlsx");
            Console.WriteLine("Workbook saved successfully with slicer positioned at TopPixel=100, LeftPixel=400.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
