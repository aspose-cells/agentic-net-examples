// Title: Aspose.Cells for .NET: Add a Pivot Table Slicer and Retrieve Its Identifier
// Description: Shows how to create a workbook, fill it with sample data, build a pivot table, add a slicer linked to the "Fruit" field, obtain the slicer's index and name (or custom name), and save the file for later API operations.
// Keywords: Aspose.Cells | C# | .NET | Excel slicer | pivot table slicer | add slicer programmatically | SlicerCollection | retrieve slicer index | retrieve slicer name | slicer identifier | Aspose.Cells API example | automate Excel slicer | workbook.Save | code snippet
// Common Searches: Aspose.Cells add slicer to pivot table C# | How to get slicer index after creating it with Aspose.Cells | Retrieve slicer name in Aspose.Cells .NET | Rename slicer programmatically Aspose.Cells | Aspose.Cells SlicerCollection example
// Developer Intent: Create a slicer linked to a pivot table and capture its index or name for subsequent API calls.
// Use Cases: Insert a slicer for a specific pivot field and store its index to modify properties later. | Assign a custom name to a newly created slicer for clearer reference in future code. | Persist slicer configuration by saving the workbook after insertion.
// AI Prompts: Generate C# code using Aspose.Cells to add a slicer for a pivot table field and return its index and name. | Show how to rename a slicer after creation and use the identifier to change its appearance or behavior. | Explain how to retrieve a slicer from a worksheet's SlicerCollection by index in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

namespace AsposeCellsSlicerDemo
{
    // Shows how to create a workbook, fill it with sample data, build a pivot table, add a slicer linked to the "Fruit" field, obtain the slicer's index and name (or custom name), and save the file for later API operations.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            sheet.Cells["A1"].PutValue("Fruit");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["A3"].PutValue("Orange");
            sheet.Cells["A4"].PutValue("Banana");
            sheet.Cells["B1"].PutValue("Quantity");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(15);

            // Add a pivot table based on the data range
            int pivotIdx = sheet.PivotTables.Add("A1:B4", "D1", "PivotTable1");
            PivotTable pivot = sheet.PivotTables[pivotIdx];
            // Use the first field (Fruit) as a row field
            pivot.AddFieldToArea(PivotFieldType.Row, 0);

            // Access the slicer collection of the worksheet
            SlicerCollection slicers = sheet.Slicers;

            // Add a slicer linked to the pivot table.
            // Parameters: pivot table, destination cell for the slicer, base field name ("Fruit")
            int slicerIdx = slicers.Add(pivot, "H1", "Fruit");

            // Retrieve the slicer object using the returned index
            Slicer slicer = slicers[slicerIdx];

            // The slicer identifier can be its index or its name.
            // Here we output both for later API usage.
            Console.WriteLine($"Slicer Index (Identifier): {slicerIdx}");
            Console.WriteLine($"Slicer Name: {slicer.Name}");

            // Optionally set a custom name for easier reference later
            slicer.Name = "FruitSlicer";

            // Save the workbook (lifecycle rule)
            workbook.Save("SlicerDemo.xlsx");
        }
    }
}
