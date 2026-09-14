// Title: How to apply the built‑in Light 1 slicer style to a pivot table slicer using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that creates a pivot table, adds a slicer linked to its row field, and applies a built‑in light slicer style using Aspose.Cells. | Show the steps to programmatically format an Excel slicer with a built‑in light theme via the Aspose.Cells API in a .NET application.
// Common Searches: Aspose.Cells C# apply Light1 slicer style to pivot table slicer | set slicer visual style to Light 1 using Aspose.Cells .NET | programmatically change slicer style in Excel with Aspose.Cells | example of using SlicerStyleLight1 in C# Aspose.Cells | how to format slicer appearance in Aspose.Cells workbook
// Tags: Aspose.Cells slicer style formatting | C# built‑in Light1 slicer style | pivot table slicer visual customization Aspose | Excel slicer appearance .NET | apply Light1 style to slicer Aspose

using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

// The example creates a workbook, adds sample data, builds a pivot table, inserts a slicer linked to the row field, sets the slicer's StyleType to the built‑in Light 1 style, and saves the file as SlicerStyleLight1.xlsx.
class ApplySlicerStyle
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data for the pivot table
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["B1"].PutValue("Value");
        worksheet.Cells["A2"].PutValue("A");
        worksheet.Cells["B2"].PutValue(10);
        worksheet.Cells["A3"].PutValue("B");
        worksheet.Cells["B3"].PutValue(20);
        worksheet.Cells["A4"].PutValue("C");
        worksheet.Cells["B4"].PutValue(30);

        // Add a pivot table based on the data range
        int pivotIndex = worksheet.PivotTables.Add("A1:B4", "D3", "PivotTable1");
        PivotTable pivotTable = worksheet.PivotTables[pivotIndex];
        pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Row field: Category
        pivotTable.AddFieldToArea(PivotFieldType.Data, 1);  // Data field: Value
        pivotTable.CalculateData();

        // Add a slicer linked to the row field (index 0) and place it at cell F3
        int slicerIndex = worksheet.Slicers.Add(pivotTable, "F3", 0);
        Slicer slicer = worksheet.Slicers[slicerIndex];

        // Apply the built‑in Light 1 slicer style
        slicer.StyleType = SlicerStyleType.SlicerStyleLight1;

        // Save the workbook with the styled slicer
        workbook.Save("SlicerStyleLight1.xlsx");
    }
}
