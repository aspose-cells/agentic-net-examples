// Title: Applying the built‑in SlicerStyleLight1 to a slicer in an Excel workbook with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells set slicer style to SlicerStyleLight1 | How to change slicer appearance using Aspose.Cells .NET API | Apply built‑in slicer style Light1 to a pivot table slicer in C# | Aspose.Cells example for formatting slicer after adding to worksheet | Saving Excel file with styled slicer using Aspose.Cells for .NET
// Tags: Aspose.Cells slicer style application | SlicerStyleLight1 usage in .NET | pivot table slicer formatting with Aspose | C# workbook save with styled slicer | set slicer StyleType property Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Slicers;
using Aspose.Cells.Pivot;

// The example creates a workbook, adds sample data, builds a pivot table, inserts a slicer linked to the 'Category' field, sets the slicer's StyleType to the built‑in SlicerStyleLight1, optionally adjusts its position and size, and saves the workbook as an .xlsx file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for a pivot table
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Product");
            sheet.Cells["C1"].PutValue("Quantity");

            sheet.Cells["A2"].PutValue("Fruits");
            sheet.Cells["B2"].PutValue("Apple");
            sheet.Cells["C2"].PutValue(10);

            sheet.Cells["A3"].PutValue("Fruits");
            sheet.Cells["B3"].PutValue("Banana");
            sheet.Cells["C3"].PutValue(15);

            sheet.Cells["A4"].PutValue("Vegetables");
            sheet.Cells["B4"].PutValue("Carrot");
            sheet.Cells["C4"].PutValue(20);

            // Create a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:C4", "E1", "PivotTable1");
            PivotTable pivot = sheet.PivotTables[pivotIndex];
            pivot.AddFieldToArea(PivotFieldType.Row, "Category");
            pivot.AddFieldToArea(PivotFieldType.Column, "Product");
            pivot.AddFieldToArea(PivotFieldType.Data, "Quantity");
            pivot.RefreshData();
            pivot.CalculateData();

            // Add a slicer linked to the pivot table for the "Category" field.
            // Destination cell (e.g., "G1") specifies where the slicer will be placed.
            int slicerIndex = sheet.Slicers.Add(pivot, "G1", "Category");
            Slicer slicer = sheet.Slicers[slicerIndex];

            // Apply the built‑in style SlicerStyleLight1 to the slicer
            slicer.StyleType = SlicerStyleType.SlicerStyleLight1;

            // Optional: set slicer position and size
            slicer.TopPixel = 100;
            slicer.LeftPixel = 400;
            slicer.WidthPixel = 150;
            slicer.HeightPixel = 200;

            // Save the workbook with the styled slicer
            string outputPath = "SlicerStyleLight1Demo.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
