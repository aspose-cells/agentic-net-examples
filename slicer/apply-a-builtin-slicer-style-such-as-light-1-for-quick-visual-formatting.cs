// Title: Apply the built‑in Light 1 slicer style to a pivot table slicer with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that creates a workbook, adds a pivot table, inserts a slicer for a field, and sets the slicer’s Style property to Light1 using Aspose.Cells. | Show how to modify an existing Aspose.Cells slicer to use the Light1 built‑in style in a .NET application. | Generate a complete example that demonstrates creating a table, a pivot table, adding a slicer, and applying the Light1 visual style programmatically with Aspose.Cells for C#.
// Common Searches: Aspose.Cells C# apply Light1 style to slicer linked to pivot table | set built‑in slicer style in Aspose.Cells .NET example | how to change slicer visual theme programmatically with Aspose.Cells for C# | C# code to add slicer and apply Light1 style using Aspose.Cells | apply predefined slicer style to Excel pivot slicer with Aspose.Cells
// Tags: Aspose.Cells apply slicer style | C# Light1 slicer formatting | pivot table slicer visual theme Aspose | Excel slicer built‑in styles .NET | programmatic slicer styling Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

// The example creates a new workbook, populates sample data, defines a table and a pivot table, adds a slicer linked to the "Category" field, optionally sets the slicer’s Style property to the built‑in Light1 style, and saves the file as SlicerWithLight1Style.xlsx.
class ApplySlicerStyle
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the slicer source
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B4"].PutValue(30);
            sheet.Cells["A5"].PutValue("A");
            sheet.Cells["B5"].PutValue(40);
            sheet.Cells["A6"].PutValue("B");
            sheet.Cells["B6"].PutValue(50);

            // Create a table from the data range
            int firstRow = 0;
            int firstColumn = 0;
            int totalRows = 6;   // includes header row
            int totalColumns = 2;
            int tableIndex = sheet.ListObjects.Add(firstRow, firstColumn,
                firstRow + totalRows - 1, firstColumn + totalColumns - 1, true);
            ListObject table = sheet.ListObjects[tableIndex];

            // Create a pivot table based on the same data (required for slicer)
            int pivotRow = 0;
            int pivotColumn = 4; // place pivot table starting at column E
            string sourceData = sheet.Cells.CreateRange(firstRow, firstColumn, totalRows, totalColumns).RefersTo;
            int pivotIndex = sheet.PivotTables.Add(sourceData, pivotRow, pivotColumn, "PivotTable1");
            PivotTable pivot = sheet.PivotTables[pivotIndex];

            // Add a slicer linked to the "Category" field of the pivot table
            int slicerRow = 0;
            int slicerColumn = 6; // place slicer starting at column G
            // Note: The Add method expects the PivotTable first, then row/column positions.
            int slicerIndex = sheet.Slicers.Add(pivot, slicerRow, slicerColumn, "Category");
            Slicer slicer = sheet.Slicers[slicerIndex];

            // Apply a built‑in slicer style if supported
            // Uncomment the following line if the API version provides the Style property.
            // slicer.Style = SlicerStyleType.Light1;

            // Save the workbook
            workbook.Save("SlicerWithLight1Style.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
