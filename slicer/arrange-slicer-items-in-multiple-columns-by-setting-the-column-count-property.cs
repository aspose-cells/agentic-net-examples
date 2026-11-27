// Title: Display slicer items in multiple columns with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells example for displaying slicer items in two columns | how to configure NumberOfColumns for an Excel slicer using Aspose.Cells .NET | set slicer column count in a pivot table with Aspose.Cells C# code | Aspose.Cells slicer multi‑column layout tutorial | create Excel slicer with multiple columns programmatically in .NET
// Tags: Aspose.Cells slicer NumberOfColumns property | C# Excel slicer multi‑column configuration | pivot table slicer layout Aspose.Cells | Aspose.Cells .NET slicer column arrangement | Excel slicer column count using Aspose API

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

// The sample creates a workbook, adds sample data and a pivot table, inserts a slicer linked to the Category field, positions it, sets the slicer’s NumberOfColumns to 2 to show items in two columns, and saves the file as SlicerMultipleColumns.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data for the pivot table
            cells["A1"].PutValue("Category");
            cells["B1"].PutValue("Product");
            cells["C1"].PutValue("Quantity");

            cells["A2"].PutValue("Fruits");
            cells["B2"].PutValue("Apple");
            cells["C2"].PutValue(10);

            cells["A3"].PutValue("Fruits");
            cells["B3"].PutValue("Banana");
            cells["C3"].PutValue(15);

            cells["A4"].PutValue("Vegetables");
            cells["B4"].PutValue("Carrot");
            cells["C4"].PutValue(20);

            cells["A5"].PutValue("Vegetables");
            cells["B5"].PutValue("Tomato");
            cells["C5"].PutValue(12);

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:C5", "E1", "PivotTable1");
            PivotTable pivot = sheet.PivotTables[pivotIndex];

            // Configure the pivot table fields
            pivot.AddFieldToArea(PivotFieldType.Row, "Category");
            pivot.AddFieldToArea(PivotFieldType.Row, "Product");
            pivot.AddFieldToArea(PivotFieldType.Data, "Quantity");
            pivot.RefreshData();   // Obsolete but still functional
            pivot.CalculateData();

            // Add a slicer linked to the "Category" field of the pivot table
            int slicerIndex = sheet.Slicers.Add(pivot, "Category", "CategorySlicer");
            Slicer slicer = sheet.Slicers[slicerIndex];

            // Basic visual settings for the slicer (using pixel‑based properties)
            slicer.Caption = "Category";
            slicer.TopPixel = 100;    // Position from the top of the worksheet (pixels)
            slicer.LeftPixel = 400;   // Position from the left of the worksheet (pixels)
            slicer.WidthPixel = 200;
            slicer.HeightPixel = 150;

            // Arrange slicer items into multiple columns
            slicer.NumberOfColumns = 2; // Display items in two columns

            // Save the workbook
            string outputPath = "SlicerMultipleColumns.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
